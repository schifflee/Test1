Option Compare Text

Imports System
Imports System.IO
Imports System.Text
'Imports Microsoft.Office.Interop.Word

Public Class frmDocumentCompare

    'small change

    Public boolHold As Boolean
    Public gDoc As String
    Public boolFormLoad As Boolean = False
    Public boolTemplate As Boolean = True
    Public gReport As String = ""
    Public strPrevForm As String = ""
    Public boolFromReportHistory As Boolean = False
    Public strDocType As String = "Final Report"
    Public gid1 As Int64
    Public gid2 As Int64
    Public gboolCompare As Boolean = False
    'Public boolExitDone As Boolean = False

    Dim da As OleDbDataAdapter = New OleDbDataAdapter()
    Dim boolSTB As Boolean = True
    Dim boolHasBeenClicked As Boolean = False
    Dim gPWD As String

    Public vTaskBar
    Public boolFromClearCompare As Boolean = False

    Public frm As Form

    Public boolHasBeenSaved = False
    'Public tempPswd As String = tPswd 'when report is from a generated report

    Public arrRanges1() As Microsoft.Office.Interop.Word.Range 'for ovDC1
    Public arrRanges2() As Microsoft.Office.Interop.Word.Range 'for ovDC2

    Public boolAsIsOp As Boolean
    Public boolReadOnlyOp As Boolean
    Public boolNoneOp As Boolean
    Public boolWaterMarkOp As Boolean
    Public boolTextOp As Boolean
    Public boolCenterOp As Boolean
    Public boolFooterOp As Boolean
    Public boolDTCreatedOp As Boolean
    Public boolDTReportedOp As Boolean
    Public boolDocIDOp As Boolean
    Public boolDocGenOp As Boolean
    Public boolDocOwnerOp As Boolean

    '20190115 LEE
    Public strPathT1 As String = ""
    Public strPathT2 As String = ""


    Sub OV_Security(ov As AxEDOfficeLib.AxEDOffice)

        Try

            ov.WordDisableCopyHotKey(True)
            'Disables the Copy keycodes in the MS Word.
            'Disable: Disables the CTRL+C, CTRL+V, CTRL+X, SHIFT+DEL, SHIFT+INSERT, ALT+CTRL+V,
            'CTRL+SHIFT+C, CTRL+INSERT

            ov.WordDisableDragAndDrop(True)
            'Disables drag and drop.

            ov.WordDisablePrintHotKey(True)
            'Disables the Print keycodes in the MS Word.
            'Disable: Disables the CTRL+P, CTRL+SHIFT+F12, CTRL+F2, ALT+CTRL+I

            ov.WordDisableSaveHotKey(True)
            'Disables the Save keycodes in the MS Word.
            'Disable: Disables the CTRL+S, SHIFT+F12, Alt+SHIFT+F2, F12

        Catch ex As Exception

        End Try

    End Sub

    Private Sub frmDocumentCompare_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed

        Dim var1

        Try
            Call ClearTemp()
        Catch ex As Exception
            var1 = var1
        End Try


    End Sub

    Private Sub frmDocumentCompare_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        '20170318 LEE: Logic when form is opened and gboolER (Enable Generated Report management.) = true

        'Document should be read-only state

        'If form load originates from a newly created report


        'If form load originates from report history


        Dim strM As String
        Dim Count1 As Int16

        strM = "LABIntegrity StudyDoc" & ChrW(8482) & " - Compare Documents"
        Me.Text = strM

        If Len(gDoc) = 0 Then
            Call ClearTemp()
        End If

        Call ControlDefaults(Me)

        Call DoubleBufferControl(Me, "dgv")

        Call DocumentCompareToolTips()

        Dim var1, var2, var3

        Call modExtensionMethods.DoubleBufferedControl(Me, True)
        Call modExtensionMethods.DoubleBufferedControl(Me.pan1, True)
        Call modExtensionMethods.DoubleBufferedControl(Me.panSC1, True)
        Try
            Call modExtensionMethods.DoubleBufferedControl(Me.ovDC, True)
        Catch ex As Exception
            var1 = ex.Message
        End Try
        Try
            Call modExtensionMethods.DoubleBufferedControl(Me.ovDC1, True)
        Catch ex As Exception
            var1 = ex.Message
        End Try
        Try
            Call modExtensionMethods.DoubleBufferedControl(Me.ovDC2, True)
        Catch ex As Exception
            var1 = ex.Message
        End Try

        Call EnableOVstuffFinalReport(Me.ovDC)
        Call EnableOVstuffFinalReport(Me.ovDC1)
        Call EnableOVstuffFinalReport(Me.ovDC2)

        If StrComp(strPrevForm, "Report History", CompareMethod.Text) = 0 Then
        Else
            boolHasBeenSaved = False
        End If

        boolFormLoad = True
        boolHold = False

        'Me.ovDC.LicenseName = "Gubbs Inc"
        'Me.ovDC.LicenseCode = "EDO8-5573-1234-ABEB" 'v8

        'Me.ovDC1.LicenseName = "Gubbs Inc"
        'Me.ovDC1.LicenseCode = "EDO8-5573-1234-ABEB" 'v8

        'Me.ovDC2.LicenseName = "Gubbs Inc"
        'Me.ovDC2.LicenseCode = "EDO8-5573-1234-ABEB" 'v8.371


        Me.ovDC.LicenseName = "LabIntegrity7631358702" 'v8.812
        Me.ovDC.LicenseCode = "EDO8-5573-1234-ABEB" 'v8.812

        Me.ovDC1.LicenseName = "LabIntegrity7631358702" 'v8.812
        Me.ovDC1.LicenseCode = "EDO8-5573-1234-ABEB" 'v8.812

        Me.ovDC2.LicenseName = "LabIntegrity7631358702" 'v8.812
        Me.ovDC2.LicenseCode = "EDO8-5573-1234-ABEB" 'v8.812

        Try
            Call EnterLabels()
        Catch ex As Exception
            MsgBox("EnterLables: " & ex.Message, vbInformation, "Problem...")
        End Try

        Call PlaceForm()

        Try
            Call PlaceControls()
        Catch ex As Exception
            MsgBox("PlaceControls: " & ex.Message, vbInformation, "Problem...")
        End Try

        Try
            Call Choice()
        Catch ex As Exception
            MsgBox("Choice: " & ex.Message, vbInformation, "Problem...")
        End Try

        'fill comboboxes
        Call FillCombos()

        Try
            Call LoadGrids()
        Catch ex As Exception
            MsgBox("LoadGrids: " & ex.Message, vbInformation, "Problem...")
        End Try

        Try
            Call ConfigGrids()
        Catch ex As Exception
            MsgBox("ConfigGrids: " & ex.Message, vbInformation, "Problem...")
        End Try


        'load textboxes
        Try
            Call LoadTextBoxes()
        Catch ex As Exception
            MsgBox("LoadTextBoxes: " & ex.Message, vbInformation, "Problem...")
        End Try

        Me.cmdWord.Enabled = True
        If gDoPDF Then
            'rearrange buttons
            Me.cmdWord.Enabled = False ' True
            Me.cmdOpenPDF.Visible = True
        Else
            Me.cmdOpenPDF.Visible = True ' False
        End If

        Try
            Call SizePanes(False)
        Catch ex As Exception

        End Try

        'load document
        If Len(gReport) = 0 Then
        Else
            Call LoadAFR(False, False, False)
        End If
        'Call LoadAFR(False, False)

        Cursor.Current = Cursors.Default

        boolFormLoad = False

        'pesky
        Try
            Call PlaceControls()
        Catch ex As Exception
            MsgBox("PlaceControls: " & ex.Message)
        End Try

        Try
            'frmH.lblProgress.Visible = False
            'frmH.pb1.Visible = False
            'frmH.pb2.Visible = False

            frmH.panProgress.Visible = False
            frmH.panProgress.Refresh()

        Catch ex As Exception

        End Try

        Call ShowButtons()

        Call PlaceLabels()

        If BOOLFINALREPORTLOCKED Then
            Me.rbFinalReport.Enabled = False
            Me.lblFinalReportLocked.Visible = True
            Me.cmdEdit.Enabled = False
        Else
            Me.rbFinalReport.Enabled = True
            Me.lblFinalReportLocked.Visible = False
            Me.cmdEdit.Enabled = True
        End If

        'pesky
        Call PlaceControls()

        Cursor.Current = Cursors.Default

        If BOOLFORCEFINALREPORTPDF And boolFromReportHistory = False Then

            Call DoPDF()

        Else

            'If BOOLEDITFINALREPORT Then
            'Else
            '    strM = "User does not have permissions to Edit reports."
            '    strM = strM & ChrW(10) & "Therefore, user may not execute a Save action."
            '    MsgBox(strM, vbInformation, "Invalid action...")
            '    GoTo end1
            'End If

        End If

        'Call DoEdit("Save") 'to set default button backcolors

        If gboolER Then

            gPWD = ""

            'Int ((6 - 1 + 1) * Rnd + 1) would return a random number between 1 and 6
            'Int ((200 - 150 + 1) * Rnd + 150) would return a random number between 150 and 200
            'Int ((999 - 100 + 1) * Rnd + 100) would return a random number between 100 and 999
            'Int ((122 - 48 + 1) * Rnd + 48) would return a random number between 48 and 122
            'For Count1 = 1 To 16
            '    var1 = Int((122 - 48 + 1) * Rnd() + 48)
            '    var2 = ChrW(var1)
            '    gPWD = gPWD & var2
            'Next

            gPWD = RandomPswd()

            ' Call DoReadOnly(Me.ovDC1)

            'Try
            '    Call DoReadOnly(Me.ovDC1)
            'Catch ex As Exception
            '    var1 = ex.Message
            '    var1 = var1
            'End Try

            'Try
            '    Dim doc As Microsoft.Office.Interop.Word.Document = Me.ovDC1.ActiveDocument

            '    If doc.ProtectionType <> WdProtectionType.wdNoProtection Then
            '        doc.Unprotect()
            '    End If

            '    doc.Protect(Password:=gPWD, NoReset:=False, Type:=Microsoft.Office.Interop.Word.WdProtectionType.wdAllowOnlyReading, UseIRM:=False, EnforceStyleLock:=False)
            '    doc.Windows(1).View.ShadeEditableRanges = False
            'Catch ex As Exception

            '    var1 = ex.Message
            '    var1 = var1

            'End Try

            If StrComp(strPrevForm, "ReportHistory", CompareMethod.Text) = 0 Then
                Call DefaultFormats("Save")
            Else
                Call SetlblInstructions()
                Call InitGeneratedReport()
            End If

            Try
                Call DoReadOnly(Me.ovDC1)
            Catch ex As Exception
                var1 = ex.Message
                var1 = var1
            End Try


        End If

end1:

        Call SpecialFormats()


    End Sub

    Sub LoadTextBoxes()

        If boolTemplate Then
            GoTo end1
        End If

        Dim str1 As String

        Try
            Dim intRow As Short
            Dim dgvR As DataGridView = frmH.dgvReports
            Dim strRN As String
            Dim strT As String
            Dim var1
            Dim strCW As String = "Loaded Document: "

            If dgvR.RowCount = 0 Then

            Else
                intRow = dgvR.CurrentCell.RowIndex

                strRN = NZ(dgvR("CHARREPORTNUMBER", intRow).Value, "")
                strT = NZ(dgvR("CHARREPORTTITLE", intRow).Value, "")

                Me.txtReportNumber.Text = wWStudyName 'strRN
                Me.txtReportTitle.Text = strT
                If boolFromReportHistory Then
                Else
                    Me.txtDescr.Text = gCHARREPORTGENERATEDSTATUS
                End If
            End If


        Catch ex As Exception

        End Try

end1:

    End Sub

    Sub FillCombos()


    End Sub

    Sub DoReadOnly(ByRef ov As AxEDOfficeLib.AxEDOffice)

        'wdAllowOnlyRevisions = 0,
        'wdAllowOnlyComments = 1,
        'wdAllowOnlyFormFields = 2,
        'wdAllowOnlyReading = 3,
        'wdNoProtection = -1,

        Dim var1
        Dim boolDA As Boolean

        Try

            '20181110 LEE:
            'hmmm. If this comes from Form Load (e.g. after a document has been generated
            'and the document has protected sections (e.g. [LOCKSECTION]
            'then get an office password message, not an error message
            Dim doc As Microsoft.Office.Interop.Word.Document
            Try

                doc = ov.ActiveDocument
                boolDA = doc.Application.DisplayAlerts
                doc.Application.DisplayAlerts = False

                ov.ProtectDoc(EDOfficeLib.WdProtectType.wdAllowOnlyReading)

                doc.Application.DisplayAlerts = boolDA
                '20180816 LEE:
                'Ensure document is in printview mode
                Try

                    Try
                        doc.Application.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
                    Catch ex As Exception
                        var1 = var1
                    End Try

                Catch ex As Exception
                    var1 = var1
                End Try

            Catch ex As Exception
                var1 = var1
            End Try
           
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

    End Sub

    Sub EnterLabels()

        Dim str1 As String
        Dim str2 As String

        'If Me.rbSbyS.Checked Then
        '    If boolTemplate Then
        '        str1 = "Choose a Report Template (Document 1)"
        '        str2 = "Choose a Report Template to compare (Document 2)"
        '    Else
        '        str1 = "Choose a Final Report version (Document 1)"
        '        str2 = "Choose a Final Report version to compare (Document 2)"
        '    End If
        'Else
        '    If boolTemplate Then
        '        str1 = "Choose a Report Template (original)"
        '        str2 = "Choose a Report Template to compare (revised)"
        '    Else
        '        str1 = "Choose a Final Report version (original)"
        '        str2 = "Choose a Final Report version to compare (revised)"
        '    End If
        'End If

        If boolTemplate Then
            str1 = "Choose a Report Template (original)"
            str2 = "Choose a Report Template to compare (revised)"
        Else
            str1 = "Choose a Final Report version (original)"
            str2 = "Choose a Final Report version to compare (revised)"
        End If

        Me.lbl1.Text = str1
        Me.lbl2.Text = str2

        str1 = "Word" & ChrW(8482) & " Compare"
        Me.rbCompare.Text = str1

    End Sub

    Sub PlaceForm()


        Dim w1, w2, h1, h2

        w1 = Screen.PrimaryScreen.WorkingArea.Width
        h1 = Screen.PrimaryScreen.WorkingArea.Height

        Me.Left = 0
        Me.Top = 0
        Me.Height = h1 ' - 10
        Me.Width = w1 ' - 20


    End Sub

    Sub PlaceControls()

        'Dim p1 As Panel = Me.pan1
        'Dim p2 As Panel = Me.pan2
        Dim pH1 As Panel = Me.panH1
        Dim pH2 As Panel = Me.panH2

        Dim sc As SplitContainer = Me.sc1


        Dim h, h1, h2, w, w1, w2, w3
        Dim l, t1, t2
        Dim l1, l2, l3

        Me.cbxFR1.Top = Me.cbxWRT1.Top
        Me.cbxFR1.Left = Me.cbxWRT1.Left

        Me.cbxFR2.Top = Me.cbxWRT2.Top
        Me.cbxFR2.Left = Me.cbxWRT2.Left

        Me.panWT.Left = Me.panStudy.Left
        Me.panWT.Top = Me.panStudy.Top


        'set pan1 height

        w = Me.Width
        h = Me.Height

        Dim bw As Int16 = (Me.Width - Me.ClientSize.Width) / 2 'form border width
        Dim tbh As Int16 = Me.Height - Me.ClientSize.Height - 2 * bw 'titlebar height

        'set pan1 left
        Me.pan1.Left = Me.panStudy.Left + Me.panStudy.Width + 3

        'set pan1 width
        Me.pan1.Width = w - Me.pan1.Left - bw - Me.panStudy.Left

        'set pan1.height
        t1 = Me.pan1.Top
        h1 = h - t1 - bw - tbh - 15
        Me.pan1.Height = h1

        'set pan1 width

        'set panSC1
        Me.panSC1.Size = Me.pan1.Size
        Me.panSC1.Location = Me.pan1.Location

        Me.panSave.Location = Me.panStudy.Location
        'Me.panSave.Size = Me.panStudy.Size
        Me.panSave.Left = Me.cmdExit.Left

        Me.panSave.Visible = Not (boolTemplate)

        Call PlaceProgress()

        Me.panProgress.Visible = False
        Me.panProgress.BringToFront()

        Me.pan1.Visible = boolTemplate
        Me.panSC1.Visible = Not (boolTemplate)

    End Sub

    Sub PlaceProgress()

        Me.panProgress.Top = Me.pan2.Top
        Me.panProgress.Left = Me.pan2.Left
        Me.panProgress.Height = Me.pan2.Height
        Me.panProgress.Width = Me.pan2.Width



        Dim str1 As String
        str1 = "Progress..."
        'Me.lblProgress.Text = str1
        Me.lblProgress.Dock = DockStyle.Fill



        'Me.panProgress1.Size = Me.sc1.Size
        'Me.panProgress1.Location = Me.sc1.Location



        'Me.panProgress1.Top = Me.sc1.Top
        'Me.panProgress1.Left = Me.sc1.Left
        'Me.panProgress1.Height = Me.sc1.Height - 20
        'Me.panProgress1.Width = Me.sc1.Width - 20

        'Me.panProgress1.Visible = False
        Me.panProgress1.BringToFront()

        'Me.lblProgress1.Text = str1
        Me.lblProgress1.Dock = DockStyle.Fill

        'Me.panProgress1.Size = Me.panSC1.Size
        Me.panProgress1.Location = Me.panSC1.Location
        Me.panProgress1.Top = Me.panSC1.Top + Me.sc1.Top + Me.lblNewDoc.Top + Me.lblNewDoc.Height + 1
        Me.panProgress1.Left = Me.panSC1.Left + 1
        Me.panProgress1.Width = Me.panSC1.Width - 2
        Me.panProgress1.Height = Me.panSC1.Height - Me.sc1.Top - (Me.lblNewDoc.Top + Me.lblNewDoc.Height) - 3


    End Sub

    Sub LoadGrids()

        Dim dgvP As DataGridView
        Dim dgvS As DataGridView
        Dim Count1 As Int32


        'dgvS = Me.dgvStudies
        'dgvS.DataSource = tblwSTUDY
        ''dgvS.DataSource = tblStudies
        'dgvS.AllowUserToAddRows = False
        'dgvS.AllowUserToDeleteRows = False

        'Call ConfigGrids()

        'For Count1 = 0 To dgvS.ColumnCount - 1
        '    dgvS.Columns(Count1).ReadOnly = True
        'Next

        'If dgvP.RowCount = 0 Then
        '    If dgvS.RowCount = 0 Then
        '        dgvS.Rows(0).Selected = True
        '    End If
        'Else
        '    dgvP.Rows(0).Selected = True
        '    Call ClickProjects()
        'End If

        If boolTemplate Then
            GoTo end1
        End If

        Dim dgvFR As DataGridView = Me.dgvFinalReports
        Dim dgvSec As DataGridView = Me.dgvSections

        Dim strF As String
        Dim strS As String

        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String

        Dim dvFR As DataView
        Dim dvSec As DataView

        strF = "CHARREPORTTYPE = 'Final Report'"
        strS = "INTFINALREPORTVERSION DESC"
        dvFR = New DataView(tblFinalReport, strF, strS, DataViewRowState.CurrentRows)
        dvFR.AllowDelete = False
        dvFR.AllowEdit = False
        dvFR.AllowNew = False
        dgvFR.SuspendLayout()
        dgvFR.DataSource = dvFR
        dgvFR.ResumeLayout()

        strF = "CHARREPORTTYPE = 'Section'"
        strS = "UPSIZE_TS DESC"
        dvSec = New DataView(tblFinalReport, strF, strS, DataViewRowState.CurrentRows)
        dvSec.AllowDelete = False
        dvSec.AllowEdit = False
        dvSec.AllowNew = False
        dgvSec.SuspendLayout()
        dgvSec.DataSource = dvSec
        dgvSec.ResumeLayout()

end1:

    End Sub

    Sub ApplyOptions()






    End Sub

    Sub ConfigGrids()

        If boolTemplate Then
            GoTo end1
        End If

        Dim Count1 As Short
        Dim Count2 As Short
        Dim str1 As String
        Dim str2 As String
        Dim intOrder As Short
        Dim var1

        Dim dgv As DataGridView

        Dim dgvFR As DataGridView = Me.dgvFinalReports


        'legend
        'ID_TBLFINALREPORT
        'ID_TBLSTUDIES
        'ID_TBLREPORTHISTORY
        'ID_TBLREPORTS
        'INTFINALREPORTVERSION
        'CHARDESCRIPTION
        'CHARCOMMENTS
        'CHARREPORTTYPE
        'BOOLLOCKED
        'ID_TBLPERSONNEL
        'ID_TBLUSERACCOUNTS
        'CHARUSERID
        'UPSIZE_TS


        For Count2 = 1 To 2
            Select Case Count2
                Case 1
                    dgv = Me.dgvFinalReports
                Case 2
                    dgv = Me.dgvSections
            End Select

            For Count1 = 0 To dgv.Columns.Count - 1
                dgv.Columns(Count1).Visible = False
            Next

            intOrder = 0

            'If Count2 = 2 Then
            intOrder = intOrder + 1
            str1 = "ID_TBLFINALREPORT"
            str2 = "ID"
            dgv.Columns(str1).HeaderText = str2
            dgv.Columns(str1).DisplayIndex = intOrder
            dgv.Columns(str1).Visible = True
            'End If

            If Count2 = 1 Then
                intOrder = intOrder + 1
                str1 = "INTFINALREPORTVERSION"
                str2 = "Ver."
                dgv.Columns(str1).HeaderText = str2
                dgv.Columns(str1).DisplayIndex = intOrder
                dgv.Columns(str1).Visible = True
            End If

            'If Count2 = 2 Then
            '    intOrder = intOrder + 1
            '    str1 = "ID_TBLFINALREPORT"
            '    str2 = "ID"
            '    dgv.Columns(str1).HeaderText = str2
            '    dgv.Columns(str1).DisplayIndex = intOrder
            '    dgv.Columns(str1).Visible = True
            'End If


            intOrder = intOrder + 1
            str1 = "CHARCOMMENTS"
            str2 = "Comments"
            dgv.Columns(str1).Visible = True
            dgv.Columns(str1).HeaderText = str2
            dgv.Columns(str1).DisplayIndex = intOrder

            intOrder = intOrder + 1
            str1 = "CHARUSERID"
            str2 = "UserID"
            dgv.Columns(str1).Visible = True
            dgv.Columns(str1).HeaderText = str2
            dgv.Columns(str1).DisplayIndex = intOrder

            intOrder = intOrder + 1
            str1 = "UPSIZE_TS"
            str2 = "Date"
            dgv.Columns(str1).Visible = True
            dgv.Columns(str1).HeaderText = str2
            dgv.Columns(str1).DisplayIndex = intOrder
            dgv.Columns(str1).DefaultCellStyle.Format = LDateFormat & " hh:mm:ss tt"


            intOrder = intOrder + 1
            str1 = "CHARDESCRIPTION"
            str2 = "Descr."
            dgv.Columns(str1).Visible = True
            dgv.Columns(str1).HeaderText = str2
            dgv.Columns(str1).DisplayIndex = intOrder


            dgv.RowHeadersWidth = 25

            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgv.AutoResizeColumns()

            'select first row
            'make sure first row in dgvVersions
            If dgv.RowCount = 0 Then
            Else
                dgv.Rows(0).Selected = True
                Try
                    dgv.CurrentCell = dgv.Item(GetVisibleCol(dgv), 0)
                Catch ex As Exception
                    var1 = var1
                End Try
            End If

            For Count1 = 0 To dgv.Columns.Count - 1
                dgv.Columns(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
            Next

        Next

end1:


    End Sub


    Sub ClickProjects()

        'Dim dgv As DataGridView
        'Dim dgvS As DataGridView

        'dgv = Me.dgvProjects
        'dgvS = Me.dgvStudies

        'Dim intRow As Int32

        'If IsNothing(dgv.CurrentRow) Then
        '    GoTo end1
        'End If

        'intRow = dgv.CurrentRow.Index

        'Dim ID As Int64
        'Dim strF As String
        'Dim strS As String
        'Dim rows() As DataRow
        'Dim tbl2 As DataTable

        'If intRow >= 0 Then
        '    ID = dgv("PROJECTID", intRow).Value
        'Else
        '    ID = 0
        'End If

        'strF = "PROJECTID = " & ID
        'Me.txtProjectID.Text = ID.ToString

        'Dim tbl1 As DataTable
        'rows = tblwSTUDY.Select(strF)
        'tbl2 = rows.CopyToDataTable
        'boolHold = True
        'dgvS.DataSource = tbl2
        'boolHold = False

        'If dgvS.RowCount = 0 Then
        '    Me.txtStudyID.Text = 0
        '    Me.txtStudyID.Text = 0
        'Else
        '    boolHold = True
        '    dgvS.Rows(0).Selected = True
        '    boolHold = False
        '    Me.txtStudyID.Text = dgvS("STUDYID", 0).Value
        'End If

end1:

    End Sub

    Sub PlaceLabels()

        Dim a, b, c, d, h, t, w, w1, w2, w3
        Dim a1, t1

        w = Me.panSC1.Width
        h = Me.panSC1.Height

        'now set sc1
        b = Me.lblComparedDoc.Top + Me.lblComparedDoc.Height + 2
        Me.sc1.Top = b
        Me.sc1.Height = h - b

    End Sub

    Sub Choice()

        Dim boolRT As Boolean

        If StrComp(gDoc, "Report Template", CompareMethod.Text) = 0 Then
            boolRT = True 'WordReportTemplate
        Else
            boolRT = False 'FinalReport
        End If

        Me.cbxWRT1.Visible = boolRT
        Me.cbxWRT2.Visible = boolRT

        Me.cbxFR1.Visible = Not (boolRT)
        Me.cbxFR2.Visible = Not (boolRT)

        Me.panStudy.Visible = False ' Not (boolRT)
        Me.panWT.Visible = boolRT

    End Sub

    Private Function p1() As Object
        Throw New NotImplementedException
    End Function

    Private Sub cmdBrowseStudy_Click(sender As Object, e As EventArgs) Handles cmdBrowseStudy.Click

        Dim boolCancel As Boolean
        Dim frmW As New frmBrowseWatson
        frmW.boolGetOracle = False
        'frmW.Text = "Retrieve a Study from the Watson" & ChrW(8482) & " Oracle database"
        frmW.Text = "Select a Study"
        frmW.boolCompare = True
        frmW.ShowDialog()

        If frmW.boolCancel Then
            frmW.Dispose()
            GoTo end1
        End If

        Dim var1

        Dim dgvPD As DataGridView = Me.dgvProjects
        Dim dgvPS As DataGridView = frmW.dgvProjects

        Dim dgvSD As DataGridView = Me.dgvStudies
        Dim dgvSS As DataGridView = frmW.dgvStudies

        Dim tblP As System.Data.DataTable
        Dim tblS As System.Data.DataTable

        tblP = dgvPS.DataSource
        tblS = dgvSS.DataSource

        Dim idP As Int64
        Dim idS As Int64

        idP = NZ(frmW.txtProjectID.Text, 0)
        idS = NZ(frmW.txtStudyID.Text, 0)

        Me.txtProjectID.Text = idP.ToString
        Me.txtStudyID.Text = idS.ToString

        Dim strFP As String
        Dim strFS As String

        Dim Count1 As Short
        ''debug
        'For Count1 = 0 To tblP.Columns.Count - 1
        '    var1 = tblP.Columns(Count1).ColumnName
        '    var1 = var1
        'Next

        'For Count1 = 0 To tblS.Columns.Count - 1
        '    var1 = tblS.Columns(Count1).ColumnName
        '    var1 = var1
        'Next

        strFP = "PROJECTID = " & idP
        strFS = "STUDYID = " & idS

        Dim dvD As DataView = New DataView(tblP, strFP, "", DataViewRowState.CurrentRows)
        Dim dvS As DataView = New DataView(tblS, strFS, "", DataViewRowState.CurrentRows)

        dgvPD.DataSource = dvD
        dgvSD.DataSource = dvS

        'unhide all columns
        For Count1 = 0 To dgvPD.Columns.Count - 1
            dgvPD.Columns(Count1).Visible = True
        Next

        For Count1 = 0 To dgvSD.Columns.Count - 1
            dgvSD.Columns(Count1).Visible = True
        Next

        Dim strP As String
        Dim strS As String
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String

        strP = ""
        If idP = 0 Then
        Else
            strP = dgvPD("PROJECTIDTEXT", 0).Value
        End If
        Me.txtProject.Text = strP

        strS = ""
        If idS = 0 Then
        Else
            str1 = "Study Name: " & NZ(dgvSD("STUDYNAME", 0).Value, "N/A")
            str2 = "Study Title: " & NZ(dgvSD("STUDYTITLE", 0).Value, "N/A")
            str3 = "Species: " & NZ(dgvSD("SPECIES", 0).Value, "N/A")
            strS = str1 & ChrW(13) & ChrW(10) & str2 & ChrW(13) & ChrW(10) & str3
        End If
        Me.txtStudy.Text = strS


        frmW.Dispose()

end1:

    End Sub

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click

        Dim intR As Short
        Dim strM As String

        If boolHasBeenClicked Then
        Else
            GoTo end2
        End If

        If Me.panSave.Visible Then
            If Me.dgvFinalReports.RowCount = 0 And Me.dgvSections.RowCount = 0 Then
                GoTo end2
            End If
        End If

        strM = "Are you sure you wish to Go Back?"
        strM = strM & ChrW(10) & ChrW(10)
        strM = strM & "Any un-saved changes will be lost."
        If boolTemplate Then
        Else

            Me.lblProgress1.Text = "Closing..."
            Me.panProgress1.Visible = True
            Me.panProgress1.Refresh()

            If BOOLFORCEFINALREPORTPDF And boolFromReportHistory = False Then
                intR = 1
            Else
                If BOOLEDITFINALREPORT Then
                    'intR = MsgBox(strM, vbOKCancel, "Go Back?")
                    intR = 1
                Else
                    intR = 1
                End If

            End If

            If intR = 1 Then
            Else
                GoTo end1
            End If
        End If

end2:

        Call ExitDoc()



end1:

    End Sub

    Sub ExitDoc()

        Dim var1

        'If boolExitDone Then
        '    'Exit Sub
        'End If

        Cursor.Current = Cursors.WaitCursor

        'Dim str1 As String
        'str1 = "Progress..." & ChrW(10) & ChrW(10)
        'str1 = str1 & "Closing down Word" & ChrW(8482) & " objects..."
        'Me.lblProgress.Text = str1
        'Me.panProgress.Visible = True
        'Me.panProgress.Refresh()

        'Me.pan2.Visible = False

        Select Case strPrevForm
            Case "Home"
                frmH.BringToFront()
                frmH.Visible = True
            Case Else
                frm.BringToFront()
                frm.Visible = True
        End Select

        Me.SendToBack()

        '20181110 LEE:
        'Why??
        'Call DoWrite()

        Try
            Me.ovDC.CloseDoc(False)
        Catch ex As Exception

        End Try

        Try
            Me.ovDC1.CloseDoc(False)
        Catch ex As Exception

        End Try

        Try
            Me.ovDC2.CloseDoc(False)
        Catch ex As Exception

        End Try

        Pause(0.25)

        'Try
        '    Me.ovDC.Dispose()
        'Catch ex As Exception

        'End Try

        'Try
        '    Me.ovDC1.Dispose()
        'Catch ex As Exception

        'End Try

        'Try
        '    Me.ovDC2.Dispose()
        'Catch ex As Exception

        'End Try


        Call ClearTemp()

        'Form is now modeless
        'need to dispose
        Try
            Me.Close()
        Catch ex As Exception
            var1 = ex.Message
        End Try

        Try
            Me.Dispose()
        Catch ex As Exception
            var1 = ex.Message
        End Try


        Cursor.Current = Cursors.Default



        'boolExitDone = True

    End Sub

    Private Sub cbxWRT1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxWRT1.SelectedIndexChanged

        If boolFormLoad Or boolHold Then
            Exit Sub
        End If

        'Call LoadAFR(Me.ovDC1)

    End Sub

    Private Sub cbxWRT2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxWRT2.SelectedIndexChanged

        If boolFormLoad Or boolHold Then
            Exit Sub
        End If

        'Call LoadAFR(Me.ovDC2)

    End Sub

    Private Sub cmdCompare_Click(sender As Object, e As EventArgs) Handles cmdCompare.Click

        If LoadCompare() Then
            'pesky
            Call SetCompareOV(2, False)
            'Call DoReviewPane()
            Me.ovDC.Visible = True
        End If

    End Sub

    Function LoadCompare() As Boolean

        Dim strF As String
        Dim intV1 As Int64
        Dim intV2 As Int64
        Dim strPath1 As String
        Dim strPath2 As String
        Dim strPathC As String
        Dim id As Int64
        Dim id1 As Int64
        Dim id2 As Int64
        Dim int1 As Int16
        Dim int2 As Int16
        Dim strM As String
        Dim str1 As String
        Dim str2 As String

        LoadCompare = False

        int1 = Me.cbxWRT1.SelectedIndex
        int2 = Me.cbxWRT2.SelectedIndex

        If int1 = -1 Then
            strM = "Please choose an Orginal version."
            MsgBox(strM, vbInformation, "Invalid action...")
            Exit Function
        End If

        If int2 = -1 Then
            strM = "Please choose a Revised version."
            MsgBox(strM, vbInformation, "Invalid action...")
            Exit Function
        End If

        Me.ovDC.Visible = False

        LoadCompare = True

        'find versions
        str1 = Me.cbxWRT1.Text
        int1 = InStr(1, str1, " ", CompareMethod.Text)
        str2 = Mid(str1, 1, int1 - 1)
        intV1 = CInt(str2)

        str1 = Me.cbxWRT2.Text
        int1 = InStr(1, str1, " ", CompareMethod.Text)
        str2 = Mid(str1, 1, int1 - 1)
        intV2 = CInt(str2)

        id = Me.txtWSID.Text

        id1 = id
        id2 = id


        strF = "ID_TBLWORDSTATEMENTS = " & id1 & " AND INTWORDVERSION = " & intV1

        Cursor.Current = Cursors.WaitCursor

        Call OpenWordDocs(id1, intV1, id2, intV2) 'this loads tblWordDocs

        Call LoadAFR(False, False, False)

        'Call ClearTemp()

        Cursor.Current = Cursors.Default
        'Me.cmdCompare.Enabled = False

    End Function

    Sub DoWrite()

        Dim var1

        Try
            Me.ovDC.UnProtectDoc()
        Catch ex As Exception
            var1 = var1
        End Try

        If Me.ovDC1.Visible Then

            Try
                Me.ovDC1.UnProtectDoc()
            Catch ex As Exception
                var1 = var1
            End Try
        End If


        ''ovDC2 is never allowed write
        'Try
        '    Me.ovDC2.UnProtectDoc()
        'Catch ex As Exception

        'End Try


    End Sub

    Sub OpenFinalReportWordDocs(id1 As Int64, intV1 As Int64, strRT1 As String, id2 As Int64, intv2 As Int64, strRT2 As String)

        Dim rs As New ADODB.Recordset
        Dim con As New ADODB.Connection
        Dim constr As String
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String

        If boolGuWuAccess Then
            constr = constrIni
        ElseIf boolGuWuSQLServer Then
            constr = constrIni ' "Provider=SQLOLEDB;" & constrIni
        ElseIf boolGuWuOracle Then
            constr = constrIniGuWuODBC
        End If

        con.Open(constr)

        If boolGuWuAccess Or boolGuWuSQLServer Or boolGuWuOracle Then
            str1 = "SELECT TBLFINALREPORTWORDDOCS.* FROM TBLFINALREPORTWORDDOCS "
            'str2 = "WHERE (ID_TBLFINALREPORTWORDDOCS = " & id1 & " AND INTFINALREPORTVERSION = " & intV1 & " AND CHARREPORTTYPE = '" & strRT1 & "') OR (ID_TBLFINALREPORTWORDDOCS = " & id2 & " AND INTFINALREPORTVERSION = " & intv2 & " AND CHARREPORTTYPE = '" & strRT2 & "')"
            str2 = "WHERE (ID_TBLFINALREPORT = " & id1 & ") OR (ID_TBLFINALREPORT = " & id2 & ")"
        End If

        strSQL = str1 & str2

        'clear tbl
        tblFinalReportWordDocs.Clear()

        rs.CursorLocation = CursorLocationEnum.adUseClient
        rs.Open(strSQL, con, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly)
        rs.ActiveConnection = Nothing
        tblFinalReportWordDocs.Clear()
        tblFinalReportWordDocs.AcceptChanges()
        tblFinalReportWordDocs.BeginLoadData()
        da.Fill(tblFinalReportWordDocs, rs)
        tblFinalReportWordDocs.EndLoadData()
        rs.Close()

        rs = Nothing

        con.Close()
        con = Nothing


    End Sub

    Sub OpenWordDocs(id1 As Int64, intV1 As Int64, id2 As Int64, intv2 As Int64)

        Dim rs As New ADODB.Recordset
        Dim con As New ADODB.Connection
        Dim constr As String
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String

        If boolGuWuAccess Then
            constr = constrIni
        ElseIf boolGuWuSQLServer Then
            constr = constrIni ' "Provider=SQLOLEDB;" & constrIni
        ElseIf boolGuWuOracle Then
            constr = constrIniGuWuODBC
        End If

        con.Open(constr)

        If boolGuWuAccess Or boolGuWuSQLServer Or boolGuWuOracle Then
            str1 = "SELECT TBLWORDDOCS.* FROM TBLWORDDOCS "
            str2 = "WHERE (ID_TBLWORDSTATEMENTS = " & id1 & " AND INTWORDVERSION = " & intV1 & ") OR (ID_TBLWORDSTATEMENTS = " & id2 & " AND INTWORDVERSION = " & intv2 & ")"
        End If

        strSQL = str1 & str2

        'clear tbl
        tblWorddocs.Clear()

        rs.CursorLocation = CursorLocationEnum.adUseClient
        rs.Open(strSQL, con, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly)
        rs.ActiveConnection = Nothing
        tblWorddocs.Clear()
        tblWorddocs.AcceptChanges()
        tblWorddocs.BeginLoadData()
        da.Fill(tblWorddocs, rs)
        tblWorddocs.EndLoadData()
        rs.Close()

        rs = Nothing

        con.Close()
        con = Nothing


    End Sub


    Function DoWordCompare() As Boolean

        Cursor.Current = Cursors.WaitCursor

        DoWordCompare = False


        Dim frm As New frmWordCompare
        frm.txtDoc1.Text = strPathT1
        frm.txtDoc2.Text = strpathT2
        frm.txtDocName1.Text = Me.txtLoadedDocDescription.Text
        frm.txtDocName2.Text = Me.txtComparedDocDescription.Text

        frm.ShowDialog()

        DoWordCompare = frm.boolCancel
        If DoWordCompare Then
            GoTo end1
        End If

end1:

        frm.Dispose()

        Cursor.Current = Cursors.Default


    End Function

    Sub LoadAFR(boolFromCompareButton As Boolean, boolSection As Boolean, boolFromCancel As Boolean)

        'look here for compatability mode stuff
        'http://office.microsoft.com/en-us/word-help/use-word-2013-to-open-documents-created-in-earlier-versions-of-word-HA102749315.aspx

        Dim var1, var2
        Dim dtbl2 As System.Data.DataTable
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
        'Dim intRow As Int16
        Dim id As Int64
        Dim intV1 As Int64
        Dim intV2 As Int64
        'Dim strPathT1 As String
        'Dim strPathT2 As String
        Dim str1 As String
        Dim str2 As String
        Dim strCW As String
        Dim strPath As String
        Dim strCFR As String
        Dim strCC As String
        Dim rows() As DataRow

        Dim boolWordCompare As Boolean = False

        Dim ov As AxEDOfficeLib.AxEDOffice ' = Me.ovDC

        Cursor.Current = Cursors.WaitCursor

        Dim boolE As Boolean
        boolE = True
        Count1 = 0
        'strPathT1 = ""
        'strPathT2 = ""

        '20140228
        'try using createxml

        Call PlaceProgress() 'pesky

        Dim strP As String
        Dim int1 As Short

        Dim boolWComp As Boolean = False

        '20190116 LEE:
        'boolTemplate = True: Coming from Word Report Template window to compare Word Report Template versions

        If boolTemplate Then

            'find versions
            str1 = Me.cbxWRT1.Text
            int1 = InStr(1, str1, " ", CompareMethod.Text)
            str2 = Mid(str1, 1, int1 - 1)
            intV1 = CInt(str2)

            str1 = Me.cbxWRT2.Text
            int1 = InStr(1, str1, " ", CompareMethod.Text)
            str2 = Mid(str1, 1, int1 - 1)
            intV2 = CInt(str2)

            'intV1 = Me.cbxWRT1.SelectedIndex + 1
            'intV2 = Me.cbxWRT2.SelectedIndex + 1

            id = Me.txtWSID.Text

            ov = Me.ovDC

            ov.Visible = False

            strP = "Progress..." & ChrW(10) & ChrW(10)
            strP = strP & "Preparing Original"
            Me.lblProgress.Text = strP
            Me.panProgress.Visible = True
            Me.panProgress.Refresh()
            Me.lblProgress.Refresh()

            strPathT1 = Createxml(id, intV1, False, False, False)

            'Legend
            'str1 = "Choose a Report Template (original)"
            'str2 = "Choose a Report Template to compare (revised)"

            Pause(0.25)
            strP = "Progress..." & ChrW(10) & ChrW(10)
            strP = strP & "Preparing Revised"
            Me.lblProgress.Text = strP
            Me.lblProgress.Refresh()

            strPathT2 = Createxml(id, intV2, True, False, True)

            Pause(0.25)

            Dim wd As New Microsoft.Office.Interop.Word.Application
            'before using ov1 to open doc
            'store this value to re-apply later
            'OfficeViewer seems to set this to False no matter what
            boolSTB = wd.ShowWindowsInTaskbar

            Try
                wd.Application.NormalTemplate.Saved = True
            Catch ex As Exception

            End Try
            Try
                wd.Quit(False)
            Catch ex As Exception
                var1 = var1
            End Try

            Try
                wd = Nothing
            Catch ex As Exception
                var1 = var1
            End Try

            Dim wdDoc As Microsoft.Office.Interop.Word.Document

            Try

                Try
                    ov.CloseDoc(False)
                Catch ex As Exception
                    var1 = var1
                End Try

                Try
                    ov.ExitOfficeApp()
                Catch ex As Exception
                    var1 = var1
                End Try

                'give the doc some time to establish, or get a read-only error
                Pause(0.5)

                strP = "Progress..." & ChrW(10) & ChrW(10)
                strP = strP & "Opening Original"
                Me.lblProgress.Text = strP
                Me.lblProgress.Refresh()

                Try
                    'ov.Open(strPathT2, "Word.Application") 'v8
                    ov.OpenWord(strPathT1)
                Catch ex As Exception
                    Try
                        ov.OpenWord(strPathT1)
                    Catch ex1 As Exception
                        MsgBox("ReportTemplate: LoadAFR (ov.OpenWord(strPathT2))):" & ChrW(10) & strPathT2 & ChrW(10) & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
                    End Try
                End Try

                Call OV_Security(ov)

                wdDoc = ov.ActiveDocument
                wdDoc.ActiveWindow.ActivePane.DisplayRulers = True
                Try
                    strP = "Progress..." & ChrW(10) & ChrW(10)
                    strP = strP & "Opening Revised"
                    Me.lblProgress.Text = strP
                    Me.lblProgress.Refresh()

                    'Dim wd1 As Microsoft.Office.Interop.Word.Application = wdDoc.Application
                    'ov.WordMergeAndCompare(strPathT2)

                Catch ex As Exception
                    var1 = ex.Message
                    MsgBox("LoadAFR: " & ex.Message)
                End Try


            Catch ex As Exception
                MsgBox("LoadAFR (Me.ov1.Open(strPathT, 'Word.Application') 'v8): " & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
            End Try

            Try
                'ov.ActiveDocument.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdNormalView
                ov.ActiveDocument.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
            Catch ex As Exception
                var1 = var1
            End Try

            'ovDC1 is always readonly at this point
            Call DoReadOnly(Me.ovDC)

end2:

        Else

            If boolFormLoad Then
                'GoTo end1
            End If

            If boolFromCancel Then 'this means need to reset window

                ov = Me.ovDC1

                strF = "ID_TBLFINALREPORT = " & gid1
                rows = tblFinalReport.Select(strF)
                id = rows(0).Item("ID_TBLFINALREPORT")
                intV1 = rows(0).Item("INTFINALREPORTVERSION") 'actually not needed
                tPswd = NZ(rows(0).Item("CHARPASSWORD"), "")
                strPath = Createxml(id, intV1, False, True, True)

                strPathT1 = strPath

                'Legend
                'str1 = "Choose a Report Template (original)"
                'str2 = "Choose a Report Template to compare (revised)"


                Dim wd As New Microsoft.Office.Interop.Word.Application

                'before using ov1 to open doc
                'store this value to re-apply later
                'OfficeViewer seems to set this to False no matter what
                Try
                    boolSTB = wd.ShowWindowsInTaskbar
                Catch ex As Exception
                    boolSTB = True
                End Try

                Try
                    wd.Application.NormalTemplate.Saved = True
                Catch ex As Exception

                End Try

                Try
                    wd.Quit(False)
                Catch ex As Exception
                    var1 = var1
                End Try

                Try
                    wd = Nothing
                Catch ex As Exception
                    var1 = var1
                End Try

                Dim wdDoc As Microsoft.Office.Interop.Word.Document

                Try

                    Try
                        ov.CloseDoc(False)
                    Catch ex As Exception
                        var1 = var1
                    End Try

                    Try
                        ov.ExitOfficeApp()
                    Catch ex As Exception
                        var1 = var1
                    End Try

                    'give the doc some time to establish, or get a read-only error
                    Pause(0.5)


                    ''cmdCancel already has a progress message
                    'strP = "Progress..." & ChrW(10) & ChrW(10)
                    'strP = strP & "Opening " & strCC
                    'Me.lblProgress1.Text = strP
                    'Me.lblProgress1.Refresh()

                    Try
                        'ov.Open(strPathT1, "Word.Application") 'v8
                        ov.OpenWord(strPathT1)
                    Catch ex As Exception
                        Try
                            'ov.Open(strPathT1, "Word.Application") 'v8
                            ov.OpenWord(strPathT1)
                        Catch ex1 As Exception
                            MsgBox("boolFromCompareButton: LoadAFR (ov.OpenWord(strPathT1))): " & ChrW(10) & strPathT1 & ChrW(10) & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
                        End Try
                    End Try

                    Call OV_Security(ov)

                    Try
                        wdDoc = ov.ActiveDocument
                    Catch ex As Exception
                        MsgBox("boolFromCompareButton: wdDoc = ov.ActiveDocument: " & ChrW(10) & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
                    End Try

                    Try
                        wdDoc.ActiveWindow.ActivePane.DisplayRulers = True
                    Catch ex As Exception
                        MsgBox("boolFromCompareButton: wdDoc.ActiveWindow.ActivePane.DisplayRulers: " & ChrW(10) & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
                    End Try

                Catch ex As Exception
                    MsgBox("boolFromCompareButton: LoadAFR (Me.ov1.Open(strPathT, 'Word.Application') 'v8): " & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
                End Try

                Try
                    ov.ActiveDocument.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
                Catch ex As Exception
                    var1 = var1
                End Try

                If ov.Visible Then
                Else
                    ov.Visible = True
                    'MsgBox("ov.visible")
                End If

            ElseIf boolFromCompareButton Then

                '20190116 LEE:
                'New logic: Word Compare is now done outside of StudyDoc because officeviewer.wordcompare function doesn't work for Word 2013 and above

                If Me.rbLoad.Checked Then
                    ov = Me.ovDC1
                    strCW = "Loaded Document: "
                    strCC = "Loaded"
                Else
                    ov = Me.ovDC2
                    strCW = "Compare With: "
                    strCC = "Compared With"
                End If

                strP = "Progress..." & ChrW(10) & ChrW(10)
                strP = strP & "Preparing " & strCC & " Document..."
                Me.lblProgress1.Text = strP
                Me.panProgress1.Visible = True
                Me.panProgress1.Refresh()
                Me.lblProgress1.Refresh()

                Dim dgv As DataGridView
                If boolSection Then
                    dgv = Me.dgvSections
                    str1 = "Section"
                Else
                    dgv = Me.dgvFinalReports
                    str1 = "Final Report"
                End If

                If dgv.RowCount = 0 Then
                    GoTo end1
                End If

                ov.Visible = False

                If Me.rbLoad.Checked Then
                    id = gid1
                Else
                    id = gid2
                End If
                strF = "ID_TBLFINALREPORT = " & id
                rows = tblFinalReport.Select(strF)

                intV1 = rows(0).Item("INTFINALREPORTVERSION") 'actually not needed

                tPswd = NZ(rows(0).Item("CHARPASSWORD"), "")

                'create label
                strCW = ReturnstrCW(id)

                If Me.rbLoad.Checked Then
                    Me.txtLoadedDocDescription.Text = strCW
                Else
                    Me.txtComparedDocDescription.Text = strCW
                End If

                'If InStr(1, strCFR, "Load", CompareMethod.Text) > 0 Then
                '    Me.lblNewDoc.Text = strCW
                'Else
                '    Me.lblCompareWith.Text = strCW
                'End If

                If Me.rbLoadCompare.Checked Then
                    strPath = Createxml(id, intV1, False, True, True)
                Else
                    strPath = Createxml(id, intV1, False, True, False)
                End If

                If Me.rbLoad.Checked Then
                    strPathT1 = strPath
                Else
                    strPathT2 = strPath
                End If


                'Legend
                'str1 = "Choose a Report Template (original)"
                'str2 = "Choose a Report Template to compare (revised)"


                Dim wd As New Microsoft.Office.Interop.Word.Application

                'before using ov1 to open doc
                'store this value to re-apply later
                'OfficeViewer seems to set this to False no matter what
                Try
                    boolSTB = wd.ShowWindowsInTaskbar
                Catch ex As Exception
                    boolSTB = True
                End Try

                Try
                    wd.Application.NormalTemplate.Saved = True
                Catch ex As Exception

                End Try

                Try
                    wd.Quit(False)
                Catch ex As Exception
                    var1 = var1
                End Try

                Try
                    wd = Nothing
                Catch ex As Exception
                    var1 = var1
                End Try

                Dim wdDoc As Microsoft.Office.Interop.Word.Document

                Try

                    Try
                        ov.CloseDoc(False)
                    Catch ex As Exception
                        var1 = var1
                    End Try

                    Try
                        ov.ExitOfficeApp()
                    Catch ex As Exception
                        var1 = var1
                    End Try

                    'give the doc some time to establish, or get a read-only error
                    Pause(0.5)


                    strP = "Progress..." & ChrW(10) & ChrW(10)
                    strP = strP & "Opening " & strCC
                    Me.lblProgress1.Text = strP
                    Me.lblProgress1.Refresh()

                    Try
                        'ov.Open(strPathT1, "Word.Application") 'v8
                        'ov.OpenWord(strPathT1)
                        ov.OpenWord(strPath)
                    Catch ex As Exception
                        Try
                            'ov.Open(strPathT1, "Word.Application") 'v8
                            'ov.OpenWord(strPathT1)
                            ov.OpenWord(strPath)
                        Catch ex1 As Exception
                            MsgBox("boolFromCompareButton: LoadAFR (ov.OpenWord(strPathT1))): " & ChrW(10) & strPathT1 & ChrW(10) & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
                        End Try
                    End Try

                    Call OV_Security(ov)

                    Try
                        wdDoc = ov.ActiveDocument
                    Catch ex As Exception
                        MsgBox("boolFromCompareButton: wdDoc = ov.ActiveDocument: " & ChrW(10) & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
                    End Try

                    Try
                        wdDoc.ActiveWindow.ActivePane.DisplayRulers = True
                    Catch ex As Exception
                        MsgBox("boolFromCompareButton: wdDoc.ActiveWindow.ActivePane.DisplayRulers: " & ChrW(10) & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
                    End Try

                Catch ex As Exception
                    MsgBox("boolFromCompareButton: LoadAFR (Me.ov1.Open(strPathT, 'Word.Application') 'v8): " & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
                End Try

                Try
                    ov.ActiveDocument.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
                Catch ex As Exception
                    var1 = var1
                End Try

                Try
                    'strip any permissions
                    Call JustRemoveProtectedRanges(ov, id)
                    Call DoReadOnly(ov)
                Catch ex As Exception
                    var1 = var1
                End Try

                'If Me.rbLoadCompare.Checked And Me.rbSbyS.Checked Then
                '    'strip any permissions
                '    Call JustRemoveProtectedRanges(ov, id)
                '    Call DoReadOnly(ov)
                'End If

                If ov.Visible Then
                Else
                    ov.Visible = True
                    'MsgBox("ov.visible")
                End If

            Else

                ov = Me.ovDC1

                If Len(gReport) = 0 Then
                    'create newest version of Final Report
                    Dim dgv As DataGridView = Me.dgvFinalReports
                    If dgv.RowCount = 0 Then
                        GoTo end1
                    End If

                    id = dgv("ID_TBLFINALREPORT", 0).Value
                    str1 = "Final Report"
                    intV1 = dgv("INTFINALREPORTVERSION", 0).Value

                    strPathT1 = Createxml(id, intV1, False, True, False)

                Else
                    strPathT1 = gReport
                End If


                Dim wd As New Microsoft.Office.Interop.Word.Application

                'before using ov1 to open doc
                'store this value to re-apply later
                'OfficeViewer seems to set this to False no matter what
                Try
                    boolSTB = wd.ShowWindowsInTaskbar
                Catch ex As Exception
                    boolSTB = True
                End Try

                Try
                    wd.Application.NormalTemplate.Saved = True
                Catch ex As Exception

                End Try

                Try
                    wd.Quit(False)
                Catch ex As Exception
                    var1 = var1
                End Try

                Try
                    wd = Nothing
                Catch ex As Exception
                    var1 = var1
                End Try

                Dim wdDoc As Microsoft.Office.Interop.Word.Document

                Try

                    Try
                        ov.CloseDoc(False)
                    Catch ex As Exception
                        var1 = var1
                    End Try

                    Try
                        ov.ExitOfficeApp()
                    Catch ex As Exception
                        var1 = var1
                    End Try

                    'give the doc some time to establish, or get a read-only error
                    Pause(0.5)

                    'first open a blank word document
                    'EDOffice.CreateNew “Word.Application”
                    'Try
                    '    ov.CreateNew("Word.Application")
                    'Catch ex As Exception
                    '    var1 = var1
                    'End Try

                    strP = "Progress..." & ChrW(10) & ChrW(10)
                    strP = strP & "Opening Original"
                    Me.lblProgress.Text = strP
                    Me.lblProgress.Refresh()

                    Try
                        'ov.Open(strPathT1, "Word.Application") 'v8
                        ov.OpenWord(strPathT1)
                    Catch ex As Exception
                        Try
                            'ov.Open(strPathT1, "Word.Application") 'v8
                            ov.OpenWord(strPathT1)
                        Catch ex1 As Exception
                            MsgBox("boolFromCompareButton: LoadAFR (ov.OpenWord(strPathT1))): " & ChrW(10) & strPathT1 & ChrW(10) & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
                        End Try
                    End Try

                    Call OV_Security(ov)

                    Try
                        wdDoc = ov.ActiveDocument
                    Catch ex As Exception
                        MsgBox("boolFromCompareButton: wdDoc = ov.ActiveDocument: " & ChrW(10) & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
                    End Try

                    Try
                        wdDoc.ActiveWindow.ActivePane.DisplayRulers = True
                    Catch ex As Exception
                        MsgBox("boolFromCompareButton: wdDoc.ActiveWindow.ActivePane.DisplayRulers: " & ChrW(10) & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
                    End Try

                Catch ex As Exception
                    MsgBox("boolFromCompareButton: LoadAFR (Me.ov1.Open(strPathT, 'Word.Application') 'v8): " & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
                End Try

                Try
                    ov.ActiveDocument.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
                Catch ex As Exception
                    var1 = var1
                End Try

            End If


        End If


end1:

        If boolWordCompare Then
            Call WordCompareFormats()
        End If

        Me.panProgress.Visible = False
        Me.panProgress.Refresh()

        'don't hide panprogress1 yet
        'Me.panProgress1.Visible = False
        'Me.panProgress1.Refresh()

        Try
            'ov.Visible = True
        Catch ex As Exception

        End Try
        Cursor.Current = Cursors.Default

    End Sub

    Sub LoadAFR_BU(boolFromCompareButton As Boolean, boolSection As Boolean, boolFromCancel As Boolean)

        'look here for compatability mode stuff
        'http://office.microsoft.com/en-us/word-help/use-word-2013-to-open-documents-created-in-earlier-versions-of-word-HA102749315.aspx

        Dim var1, var2
        Dim dtbl2 As System.Data.DataTable
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
        'Dim intRow As Int16
        Dim id As Int64
        Dim intV1 As Int64
        Dim intV2 As Int64
        'Dim strPathT1 As String
        'Dim strPathT2 As String
        Dim str1 As String
        Dim str2 As String
        Dim strCW As String
        Dim strPath As String
        Dim strCFR As String
        Dim strCC As String
        Dim rows() As DataRow

        Dim boolWordCompare As Boolean = False

        Dim ov As AxEDOfficeLib.AxEDOffice ' = Me.ovDC

        Cursor.Current = Cursors.WaitCursor

        Dim boolE As Boolean
        boolE = True
        Count1 = 0
        'strPathT1 = ""
        'strPathT2 = ""

        '20140228
        'try using createxml

        Call PlaceProgress() 'pesky

        Dim strP As String
        Dim int1 As Short

        Dim boolWComp As Boolean = False

        If boolTemplate Then

            'find versions
            str1 = Me.cbxWRT1.Text
            int1 = InStr(1, str1, " ", CompareMethod.Text)
            str2 = Mid(str1, 1, int1 - 1)
            intV1 = CInt(str2)

            str1 = Me.cbxWRT2.Text
            int1 = InStr(1, str1, " ", CompareMethod.Text)
            str2 = Mid(str1, 1, int1 - 1)
            intV2 = CInt(str2)

            'intV1 = Me.cbxWRT1.SelectedIndex + 1
            'intV2 = Me.cbxWRT2.SelectedIndex + 1

            id = Me.txtWSID.Text

            ov = Me.ovDC

            ov.Visible = False

            strP = "Progress..." & ChrW(10) & ChrW(10)
            strP = strP & "Preparing Original"
            Me.lblProgress.Text = strP
            Me.panProgress.Visible = True
            Me.panProgress.Refresh()
            Me.lblProgress.Refresh()

            strPathT1 = Createxml(id, intV1, False, False, False)

            'Legend
            'str1 = "Choose a Report Template (original)"
            'str2 = "Choose a Report Template to compare (revised)"

            Pause(0.25)
            strP = "Progress..." & ChrW(10) & ChrW(10)
            strP = strP & "Preparing Revised"
            Me.lblProgress.Text = strP
            Me.lblProgress.Refresh()

            strPathT2 = Createxml(id, intV2, True, False, True)

            Pause(0.25)

            Dim wd As New Microsoft.Office.Interop.Word.Application
            'before using ov1 to open doc
            'store this value to re-apply later
            'OfficeViewer seems to set this to False no matter what
            boolSTB = wd.ShowWindowsInTaskbar

            Try
                wd.Application.NormalTemplate.Saved = True
            Catch ex As Exception

            End Try
            Try
                wd.Quit(False)
            Catch ex As Exception
                var1 = var1
            End Try

            Try
                wd = Nothing
            Catch ex As Exception
                var1 = var1
            End Try

            Dim wdDoc As Microsoft.Office.Interop.Word.Document

            Try

                Try
                    ov.CloseDoc(False)
                Catch ex As Exception
                    var1 = var1
                End Try

                Try
                    ov.ExitOfficeApp()
                Catch ex As Exception
                    var1 = var1
                End Try

                'give the doc some time to establish, or get a read-only error
                Pause(0.5)

                strP = "Progress..." & ChrW(10) & ChrW(10)
                strP = strP & "Opening Original"
                Me.lblProgress.Text = strP
                Me.lblProgress.Refresh()

                Try
                    'ov.Open(strPathT2, "Word.Application") 'v8
                    ov.OpenWord(strPathT1)
                Catch ex As Exception
                    Try
                        ov.OpenWord(strPathT1)
                    Catch ex1 As Exception
                        MsgBox("ReportTemplate: LoadAFR (ov.OpenWord(strPathT2))):" & ChrW(10) & strPathT2 & ChrW(10) & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
                    End Try
                End Try

                Call OV_Security(ov)

                wdDoc = ov.ActiveDocument
                wdDoc.ActiveWindow.ActivePane.DisplayRulers = True
                Try
                    strP = "Progress..." & ChrW(10) & ChrW(10)
                    strP = strP & "Opening Revised"
                    Me.lblProgress.Text = strP
                    Me.lblProgress.Refresh()

                    'Dim wd1 As Microsoft.Office.Interop.Word.Application = wdDoc.Application
                    'ov.WordMergeAndCompare(strPathT2)

                    '20190114 LEE:
                    'OfficeViewer.WordMergeAndCompare does not work with Word 2013 or greater
                    'have to do it through Word
                    '20190116 LEE:
                    'don't do it here though
                    'boolWComp = DoWordCompare()
                    'If boolWComp Then
                    '    'quit
                    '    GoTo end2
                    'End If

                    'Pause(0.25)

                    'Call SetCompareOV(1, True)

                Catch ex As Exception
                    var1 = ex.Message
                    MsgBox("LoadAFR: " & ex.Message)
                End Try


            Catch ex As Exception
                MsgBox("LoadAFR (Me.ov1.Open(strPathT, 'Word.Application') 'v8): " & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
            End Try

            Try
                'ov.ActiveDocument.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdNormalView
                ov.ActiveDocument.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
            Catch ex As Exception
                var1 = var1
            End Try

            'ovDC1 is always readonly at this point
            Call DoReadOnly(Me.ovDC)

end2:

        Else

            If boolFormLoad Then
                'GoTo end1
            End If

            If boolFromCancel Then

                ov = Me.ovDC1

                strF = "ID_TBLFINALREPORT = " & gid1
                rows = tblFinalReport.Select(strF)
                id = rows(0).Item("ID_TBLFINALREPORT")
                intV1 = rows(0).Item("INTFINALREPORTVERSION") 'actually not needed
                tPswd = NZ(rows(0).Item("CHARPASSWORD"), "")
                strPath = Createxml(id, intV1, False, True, True)

                strPathT1 = strPath

                'Legend
                'str1 = "Choose a Report Template (original)"
                'str2 = "Choose a Report Template to compare (revised)"


                Dim wd As New Microsoft.Office.Interop.Word.Application

                'before using ov1 to open doc
                'store this value to re-apply later
                'OfficeViewer seems to set this to False no matter what
                Try
                    boolSTB = wd.ShowWindowsInTaskbar
                Catch ex As Exception
                    boolSTB = True
                End Try

                Try
                    wd.Application.NormalTemplate.Saved = True
                Catch ex As Exception

                End Try

                Try
                    wd.Quit(False)
                Catch ex As Exception
                    var1 = var1
                End Try

                Try
                    wd = Nothing
                Catch ex As Exception
                    var1 = var1
                End Try

                Dim wdDoc As Microsoft.Office.Interop.Word.Document

                Try

                    Try
                        ov.CloseDoc(False)
                    Catch ex As Exception
                        var1 = var1
                    End Try

                    Try
                        ov.ExitOfficeApp()
                    Catch ex As Exception
                        var1 = var1
                    End Try

                    'give the doc some time to establish, or get a read-only error
                    Pause(0.5)


                    ''cmdCancel already has a progress message
                    'strP = "Progress..." & ChrW(10) & ChrW(10)
                    'strP = strP & "Opening " & strCC
                    'Me.lblProgress1.Text = strP
                    'Me.lblProgress1.Refresh()

                    Try
                        'ov.Open(strPathT1, "Word.Application") 'v8
                        ov.OpenWord(strPathT1)
                    Catch ex As Exception
                        Try
                            'ov.Open(strPathT1, "Word.Application") 'v8
                            ov.OpenWord(strPathT1)
                        Catch ex1 As Exception
                            MsgBox("boolFromCompareButton: LoadAFR (ov.OpenWord(strPathT1))): " & ChrW(10) & strPathT1 & ChrW(10) & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
                        End Try
                    End Try

                    Call OV_Security(ov)

                    Try
                        wdDoc = ov.ActiveDocument
                    Catch ex As Exception
                        MsgBox("boolFromCompareButton: wdDoc = ov.ActiveDocument: " & ChrW(10) & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
                    End Try

                    Try
                        wdDoc.ActiveWindow.ActivePane.DisplayRulers = True
                    Catch ex As Exception
                        MsgBox("boolFromCompareButton: wdDoc.ActiveWindow.ActivePane.DisplayRulers: " & ChrW(10) & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
                    End Try

                Catch ex As Exception
                    MsgBox("boolFromCompareButton: LoadAFR (Me.ov1.Open(strPathT, 'Word.Application') 'v8): " & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
                End Try

                Try
                    ov.ActiveDocument.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
                Catch ex As Exception
                    var1 = var1
                End Try

                If ov.Visible Then
                Else
                    ov.Visible = True
                    'MsgBox("ov.visible")
                End If

            ElseIf boolFromCompareButton Then


                If Me.rbLoad.Checked Or (Me.rbLoadCompare.Checked And Me.rbSbyS.Checked) Then

                    strCFR = Me.cmdCompareFinalReport.Text
                    If InStr(1, strCFR, "Load", CompareMethod.Text) > 0 Then
                        ov = Me.ovDC1
                        strCW = "Loaded Document: "
                        strCC = "Loaded"
                    Else
                        ov = Me.ovDC2
                        strCW = "Compare With: "
                        strCC = "Compared With"
                    End If

                    strP = "Progress..." & ChrW(10) & ChrW(10)
                    strP = strP & "Preparing " & strCC & " Document..."
                    Me.lblProgress1.Text = strP
                    Me.panProgress1.Visible = True
                    Me.panProgress1.Refresh()
                    Me.lblProgress1.Refresh()

                    Dim dgv As DataGridView
                    If boolSection Then
                        dgv = Me.dgvSections
                        str1 = "Section"
                    Else
                        dgv = Me.dgvFinalReports
                        str1 = "Final Report"
                    End If

                    If dgv.RowCount = 0 Then
                        GoTo end1
                    End If

                    ov.Visible = False

                    If Me.rbLoad.Checked Then
                        id = gid1
                    Else
                        id = gid2
                    End If
                    strF = "ID_TBLFINALREPORT = " & id
                    rows = tblFinalReport.Select(strF)


                    'Dim strDesc As String
                    'strDesc = rows(0).Item("CHARDESCRIPTION")
                    'Me.txtDescr.Text = strDesc

                    'Dim strDT As String
                    'strDT = dgv("CHARREPORTTYPE", intRow).Value
                    'Me.txtDocType.Text = strDT

                    intV1 = rows(0).Item("INTFINALREPORTVERSION") 'actually not needed

                    'If InStr(1, strCFR, "Load", CompareMethod.Text) > 0 Then
                    '    gid1 = rows(0).Item("ID_TBLFINALREPORT", intRow).Value
                    'Else
                    '    gid2 = rows(0).Item("ID_TBLFINALREPORT", intRow).Value
                    'End If

                    tPswd = NZ(rows(0).Item("CHARPASSWORD"), "")


                    'create label
                    strCW = ReturnstrCW(id)

                    If Me.rbLoad.Checked Then
                        Me.txtLoadedDocDescription.Text = strCW
                    Else
                        Me.txtComparedDocDescription.Text = strCW
                    End If

                    'If InStr(1, strCFR, "Load", CompareMethod.Text) > 0 Then
                    '    Me.lblNewDoc.Text = strCW
                    'Else
                    '    Me.lblCompareWith.Text = strCW
                    'End If

                    If Me.rbLoadCompare.Checked Then
                        strPath = Createxml(id, intV1, False, True, True)
                    Else
                        strPath = Createxml(id, intV1, False, True, False)
                    End If

                    If Me.rbLoad.Checked Then
                        strPathT1 = strPath
                    Else
                        strPathT2 = strPath
                    End If


                    'Legend
                    'str1 = "Choose a Report Template (original)"
                    'str2 = "Choose a Report Template to compare (revised)"


                    Dim wd As New Microsoft.Office.Interop.Word.Application

                    'before using ov1 to open doc
                    'store this value to re-apply later
                    'OfficeViewer seems to set this to False no matter what
                    Try
                        boolSTB = wd.ShowWindowsInTaskbar
                    Catch ex As Exception
                        boolSTB = True
                    End Try

                    Try
                        wd.Application.NormalTemplate.Saved = True
                    Catch ex As Exception

                    End Try

                    Try
                        wd.Quit(False)
                    Catch ex As Exception
                        var1 = var1
                    End Try

                    Try
                        wd = Nothing
                    Catch ex As Exception
                        var1 = var1
                    End Try

                    Dim wdDoc As Microsoft.Office.Interop.Word.Document

                    Try

                        Try
                            ov.CloseDoc(False)
                        Catch ex As Exception
                            var1 = var1
                        End Try

                        Try
                            ov.ExitOfficeApp()
                        Catch ex As Exception
                            var1 = var1
                        End Try

                        'give the doc some time to establish, or get a read-only error
                        Pause(0.5)


                        strP = "Progress..." & ChrW(10) & ChrW(10)
                        strP = strP & "Opening " & strCC
                        Me.lblProgress1.Text = strP
                        Me.lblProgress1.Refresh()

                        Try
                            'ov.Open(strPathT1, "Word.Application") 'v8
                            'ov.OpenWord(strPathT1)
                            ov.OpenWord(strPath)
                        Catch ex As Exception
                            Try
                                'ov.Open(strPathT1, "Word.Application") 'v8
                                'ov.OpenWord(strPathT1)
                                ov.OpenWord(strPath)
                            Catch ex1 As Exception
                                MsgBox("boolFromCompareButton: LoadAFR (ov.OpenWord(strPathT1))): " & ChrW(10) & strPathT1 & ChrW(10) & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
                            End Try
                        End Try

                        Call OV_Security(ov)

                        Try
                            wdDoc = ov.ActiveDocument
                        Catch ex As Exception
                            MsgBox("boolFromCompareButton: wdDoc = ov.ActiveDocument: " & ChrW(10) & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
                        End Try

                        Try
                            wdDoc.ActiveWindow.ActivePane.DisplayRulers = True
                        Catch ex As Exception
                            MsgBox("boolFromCompareButton: wdDoc.ActiveWindow.ActivePane.DisplayRulers: " & ChrW(10) & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
                        End Try

                    Catch ex As Exception
                        MsgBox("boolFromCompareButton: LoadAFR (Me.ov1.Open(strPathT, 'Word.Application') 'v8): " & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
                    End Try

                    Try
                        ov.ActiveDocument.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
                    Catch ex As Exception
                        var1 = var1
                    End Try

                    If Me.rbLoadCompare.Checked And Me.rbSbyS.Checked Then

                        'strip any permissions
                        Call JustRemoveProtectedRanges(ov, id)
                        Call DoReadOnly(ov)

                    End If

                    If ov.Visible Then
                    Else
                        ov.Visible = True
                        'MsgBox("ov.visible")
                    End If


                Else 'Word Compare

                    boolWordCompare = True

                    Try
                        Me.sc1.Panel2Collapsed = True
                    Catch ex As Exception

                    End Try


                    intV1 = 0 ' Me.cbxWRT1.SelectedIndex + 1
                    intV2 = 0 ' Me.cbxWRT2.SelectedIndex + 1
                    id = gid2 ' remember, these have to opened backwards' Me.txtWSID.Text

                    ov = Me.ovDC1

                    ov.Visible = False

                    strP = "Progress..." & ChrW(10) & ChrW(10)
                    strP = strP & "Preparing Original"
                    Me.lblProgress1.Text = strP
                    Me.panProgress1.Visible = True
                    Me.panProgress1.Refresh()
                    Me.lblProgress1.Refresh()

                    Dim dgv As DataGridView
                    If boolSection Then
                        dgv = Me.dgvSections
                        str1 = "Section"
                    Else
                        dgv = Me.dgvFinalReports
                        str1 = "Final Report"
                    End If

                    'create label
                    strCW = ReturnstrCW(id)

                    If Me.rbLoad.Checked Then
                        Me.txtLoadedDocDescription.Text = strCW
                    Else
                        Me.txtComparedDocDescription.Text = strCW
                    End If

                    If InStr(1, strCFR, "Load", CompareMethod.Text) > 0 Then
                        Me.lblNewDoc.Text = strCW
                    Else
                        Me.lblCompareWith.Text = strCW
                    End If

                    strPathT1 = Createxml(id, intV1, False, False, False)

                    'Legend
                    'str1 = "Choose a Report Template (original)"
                    'str2 = "Choose a Report Template to compare (revised)"

                    'assign to ov right away or next createxml will delete it
                    'MsgBox("strPathT1: " & strPathT1)
                    Pause(0.25)
                    strP = "Progress..." & ChrW(10) & ChrW(10)
                    strP = strP & "Preparing Revised"
                    Me.lblProgress1.Text = strP
                    Me.panProgress1.Refresh()
                    Me.lblProgress1.Refresh()

                    strPathT2 = Createxml(gid1, intV2, True, False, True)
                    'MsgBox("strPathT2: " & strPathT2)
                    Pause(0.25)

                    Dim wd As New Microsoft.Office.Interop.Word.Application
                    'before using ov1 to open doc
                    'store this value to re-apply later
                    'OfficeViewer seems to set this to False no matter what
                    boolSTB = wd.ShowWindowsInTaskbar

                    Try
                        wd.Application.NormalTemplate.Saved = True
                    Catch ex As Exception

                    End Try

                    Try
                        wd.Quit(False)
                    Catch ex As Exception
                        var1 = var1
                    End Try

                    Try
                        wd = Nothing
                    Catch ex As Exception
                        var1 = var1
                    End Try

                    Dim wdDoc As Microsoft.Office.Interop.Word.Document

                    Try

                        Try
                            ov.CloseDoc(False)
                        Catch ex As Exception
                            var1 = var1
                        End Try

                        Try
                            ov.ExitOfficeApp()
                        Catch ex As Exception
                            var1 = var1
                        End Try

                        'give the doc some time to establish, or get a read-only error
                        Pause(0.5)

                        strP = "Progress..." & ChrW(10) & ChrW(10)
                        strP = strP & "Opening Original"
                        Me.lblProgress1.Text = strP
                        Me.panProgress1.Refresh()
                        Me.lblProgress1.Refresh()

                        Try
                            'ov.Open(strPathT2, "Word.Application") 'v8
                            ov.OpenWord(strPathT2)
                        Catch ex As Exception
                            Try
                                ov.OpenWord(strPathT2)
                            Catch ex1 As Exception
                                MsgBox("LoadAFR (ov.OpenWord(strPathT2))): " & ChrW(10) & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
                            End Try
                        End Try

                        Call OV_Security(ov)

                        'must check if document has protection
                        'if so, must strip
                        'check for password
                        strF = "ID_TBLFINALREPORT = " & gid1
                        rows = tblFinalReport.Select(strF)
                        var1 = NZ(rows(0).Item("CHARPASSWORD"), "")
                        If Len(var1) = 0 Then
                        Else
                            Call JustRemoveProtectedRanges(ov, gid1)
                        End If

                        wdDoc = ov.ActiveDocument
                        wdDoc.ActiveWindow.ActivePane.DisplayRulers = True
                        Try
                            strP = "Progress..." & ChrW(10) & ChrW(10)
                            strP = strP & "Opening Revised"
                            Me.lblProgress.Text = strP
                            Me.lblProgress.Refresh()

                            'Dim wd As Microsoft.Office.Interop.Word.Application = wdDoc.Application

                            'must check if document has protection
                            'if so, must strip
                            'check for password
                            strF = "ID_TBLFINALREPORT = " & id
                            rows = tblFinalReport.Select(strF)
                            var1 = NZ(rows(0).Item("CHARPASSWORD"), "")
                            If Len(var1) = 0 Then
                            Else
                                'apply to ovd2
                                Try
                                    Me.ovDC2.CloseDoc(False)
                                Catch ex As Exception

                                End Try
                                Me.ovDC2.OpenWord(strPathT1)

                                Call JustRemoveProtectedRanges(Me.ovDC2, id)
                                'now save the document
                                Me.ovDC2.Save()
                                Try
                                    Me.ovDC2.CloseDoc(False)
                                    Pause(0.25)
                                Catch ex As Exception

                                End Try

                            End If

                            'ov.WordMergeAndCompare(strPathT1)

                            '20190114 LEE:
                            'OfficeViewer.WordMergeAndCompare does not work with Word 2013 or greater
                            'have to do it through Word

                            var1 = var1
                            'boolWComp = DoWordCompare()
                            'If boolWComp Then
                            '    'quit
                            '    'GoTo end2
                            'End If

                            'Pause(0.25)

                            'Call DoReadOnly(Me.ovDC1)

                            'Call SetCompareOV(1, True)

                        Catch ex As Exception
                            var1 = ex.Message
                            MsgBox("LoadAFR: " & ex.Message)
                        End Try


                    Catch ex As Exception
                        MsgBox("LoadAFR (Me.ov1.Open(strPathT, 'Word.Application') 'v8): " & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
                    End Try

                    Try
                        'ov.ActiveDocument.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdNormalView
                        ov.ActiveDocument.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView

                    Catch ex As Exception
                        var1 = var1
                    End Try

                End If

            Else

                ov = Me.ovDC1

                If Len(gReport) = 0 Then
                    'create newest version of Final Report
                    Dim dgv As DataGridView = Me.dgvFinalReports
                    If dgv.RowCount = 0 Then
                        GoTo end1
                    End If

                    id = dgv("ID_TBLFINALREPORT", 0).Value
                    str1 = "Final Report"
                    intV1 = dgv("INTFINALREPORTVERSION", 0).Value

                    strPathT1 = Createxml(id, intV1, False, True, False)

                Else
                    strPathT1 = gReport
                End If


                Dim wd As New Microsoft.Office.Interop.Word.Application

                'before using ov1 to open doc
                'store this value to re-apply later
                'OfficeViewer seems to set this to False no matter what
                Try
                    boolSTB = wd.ShowWindowsInTaskbar
                Catch ex As Exception
                    boolSTB = True
                End Try

                Try
                    wd.Application.NormalTemplate.Saved = True
                Catch ex As Exception

                End Try

                Try
                    wd.Quit(False)
                Catch ex As Exception
                    var1 = var1
                End Try

                Try
                    wd = Nothing
                Catch ex As Exception
                    var1 = var1
                End Try

                Dim wdDoc As Microsoft.Office.Interop.Word.Document

                Try

                    Try
                        ov.CloseDoc(False)
                    Catch ex As Exception
                        var1 = var1
                    End Try

                    Try
                        ov.ExitOfficeApp()
                    Catch ex As Exception
                        var1 = var1
                    End Try

                    'give the doc some time to establish, or get a read-only error
                    Pause(0.5)

                    'first open a blank word document
                    'EDOffice.CreateNew “Word.Application”
                    'Try
                    '    ov.CreateNew("Word.Application")
                    'Catch ex As Exception
                    '    var1 = var1
                    'End Try

                    strP = "Progress..." & ChrW(10) & ChrW(10)
                    strP = strP & "Opening Original"
                    Me.lblProgress.Text = strP
                    Me.lblProgress.Refresh()

                    Try
                        'ov.Open(strPathT1, "Word.Application") 'v8
                        ov.OpenWord(strPathT1)
                    Catch ex As Exception
                        Try
                            'ov.Open(strPathT1, "Word.Application") 'v8
                            ov.OpenWord(strPathT1)
                        Catch ex1 As Exception
                            MsgBox("boolFromCompareButton: LoadAFR (ov.OpenWord(strPathT1))): " & ChrW(10) & strPathT1 & ChrW(10) & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
                        End Try
                    End Try

                    Call OV_Security(ov)

                    Try
                        wdDoc = ov.ActiveDocument
                    Catch ex As Exception
                        MsgBox("boolFromCompareButton: wdDoc = ov.ActiveDocument: " & ChrW(10) & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
                    End Try

                    Try
                        wdDoc.ActiveWindow.ActivePane.DisplayRulers = True
                    Catch ex As Exception
                        MsgBox("boolFromCompareButton: wdDoc.ActiveWindow.ActivePane.DisplayRulers: " & ChrW(10) & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
                    End Try

                Catch ex As Exception
                    MsgBox("boolFromCompareButton: LoadAFR (Me.ov1.Open(strPathT, 'Word.Application') 'v8): " & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
                End Try

                Try
                    ov.ActiveDocument.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
                Catch ex As Exception
                    var1 = var1
                End Try

            End If


        End If


end1:

        If boolWordCompare Then
            Call WordCompareFormats()
        End If

        Me.panProgress.Visible = False
        Me.panProgress.Refresh()

        'don't hide panprogress1 yet
        'Me.panProgress1.Visible = False
        'Me.panProgress1.Refresh()

        Try
            'ov.Visible = True
        Catch ex As Exception

        End Try
        Cursor.Current = Cursors.Default

    End Sub

    Function ReturnstrCW(id As Int64) As String

        Dim var1
        Dim strCW As String

        Dim strF As String = "ID_TBLFINALREPORT = " & id
        Dim rows() As DataRow = tblFinalReport.Select(strF)

        Dim str1 As String

        str1 = NZ(rows(0).Item("CHARREPORTTYPE"), "Sample Analysis")

        Dim boolSection As Boolean
        If StrComp(str1, "Section", CompareMethod.Text) = 0 Then
            boolSection = True
        Else
            boolSection = False
        End If

        If boolSection Then

            var1 = "Section/Draft: "
            strCW = var1
            var1 = " ID: " & rows(0).Item("ID_TBLFINALREPORT")
            strCW = strCW & var1
            var1 = ", Comments: " & rows(0).Item("CHARCOMMENTS")
            strCW = strCW & var1
            var1 = ", UserID: " & rows(0).Item("CHARUSERID")
            strCW = strCW & var1
            '20181112 LEE:
            'Too much info
            'var1 = ", Date Created: " & Format(rows(0).Item("UPSIZE_TS"), "dd-MMM-yyyy hh:mm:ss tt")
            'strCW = strCW & var1
            'var1 = ", Descr: " & rows(0).Item("CHARDESCRIPTION")
            'strCW = strCW & var1
        Else

            var1 = "Final Report: "
            strCW = var1
            var1 = " ID: " & rows(0).Item("ID_TBLFINALREPORT")
            strCW = strCW & var1
            var1 = ", Ver: " & rows(0).Item("INTFINALREPORTVERSION")
            strCW = strCW & var1
            var1 = ", Comments: " & rows(0).Item("CHARCOMMENTS")
            strCW = strCW & var1
            var1 = ", UserID: " & rows(0).Item("CHARUSERID")
            strCW = strCW & var1
            '20181112 LEE:
            'Too much info
            'var1 = ", Date Created: " & Format(rows(0).Item("UPSIZE_TS"), "dd-MMM-yyyy hh:mm:ss tt")
            'strCW = strCW & var1
            'var1 = ", Descr: " & rows(0).Item("CHARDESCRIPTION")
            'strCW = strCW & var1

        End If

        ReturnstrCW = strCW

    End Function

    Function Createxml(id As Int64, intVersion As Int64, boolTempReport As Boolean, boolReport As Boolean, boolCompare As Boolean) As String

        'save as temp then display in afr
        Dim strP As String

        If boolCompare Then
            strP = GetNewTempFileReport(False)
        Else
            strP = GetNewTempFile(False)
        End If

        strP = Replace(strP, ".xml", ".docx", 1, -1, CompareMethod.Text)

        Dim intRow As Short
        Dim strPath As String
        'Dim dgv As DataGridView
        Dim strLbl As String


        Dim var1, var2
        Dim dtbl1 As System.Data.DataTable
        Dim dtbl2 As System.Data.DataTable
        Dim rows1() As DataRow
        Dim rows2() As DataRow
        Dim strF As String
        Dim strS As String
        Dim intL As Int64
        Dim strpathT As String
        Dim Count1 As Int16
        Dim strW As String
        Dim fs As FileStream
        Dim strM As String
        'Dim id As Int64
        'Dim intVersion As Int64


        If boolTemplate Then
            dtbl1 = tblWordStatements
            dtbl2 = tblWorddocs
            strF = "ID_TBLWORDSTATEMENTS = " & id & " AND INTWORDVERSION = " & intVersion
            strS = "ID_TBLWORDDOCS ASC"
        Else
            Try
                Call OpenFinalReportWordDocs(id, 0, "", 0, 0, "")
            Catch ex As Exception
                var1 = var1
            End Try

            dtbl1 = tblFinalReport
            dtbl2 = tblFinalReportWordDocs
            If boolReport Then
                strF = "ID_TBLFINALREPORT = " & id ' & " AND INTFINALREPORTVERSION = " & intVersion & " AND CHARREPORTTYPE = 'Final Report'"
            Else
                strF = "ID_TBLFINALREPORT = " & id ' & " AND CHARREPORTTYPE = 'Section'"
            End If
            strS = "ID_TBLFINALREPORTWORDDOCS ASC"
        End If

        Try
            rows2 = dtbl2.Select(strF, strS)
        Catch ex As Exception
            var1 = var1
        End Try

        intL = rows2.Length

        Dim boolE As Boolean
        boolE = True
        Count1 = 0
        strpathT = ""
        If boolTempReport Then
            strpathT = GetNewTempFileReport(False)
        Else
            strpathT = GetNewTempFile(False)
        End If

        'why am i putting a .docx on it?
        'why not leave it as xml?
        'especially since the file is being built as xml
        'leave it as xml
        '20180125 LEE: Also because document may have VBA project, so must choose docx or docm
        'make this evaluation later
        'strpathT = Replace(strpathT, ".xml", ".docx", 1, -1, CompareMethod.Text)

        strW = ""
        Dim strBuild As New StringBuilder("")

        For Count1 = 0 To intL - 1
            strBuild.Append(rows2(Count1).Item("CHARXML"))
        Next
        strW = strBuild.ToString()

        ' Add some information to the file.
        Dim info As Byte()
        If intL = 0 Then
            strM = "There is a problem with this data:" & ChrW(10)
            strM = strM & "tblWorddocs: " & strF & ChrW(10)
            strM = strM & "Please contact your StudyDoc system administrator."
            info = New UTF8Encoding(True).GetBytes(strM)
            strpathT = Replace(strpathT, ".XML", ".TXT", 1, -1, CompareMethod.Text)
        Else
            ' Add some information to the file.
            info = New UTF8Encoding(True).GetBytes(strW)
        End If

        Try
            fs = File.Create(strpathT)
            fs.Close()
            fs = File.OpenWrite(strpathT)

            fs.Write(info, 0, info.Length)
            fs.Close()
        Catch ex As Exception
            var1 = var1
        End Try

        Createxml = strpathT

        ''now convert to .docx
        ''must find if it is a docx or docm

        'Dim strPathA As String
        'Dim strExt As String
        'Dim strR As String
        'Dim wd As New Microsoft.Office.Interop.Word.Application

        'var1 = wd.Version

        'Dim wdDoc As Microsoft.Office.Interop.Word.Document
        'wdDoc = wd.Documents.Open(strpathT)
        'If wdDoc.HasVBProject Then
        '    strExt = ".docm"
        '    strR = Replace(strpathT, ".xml", strExt, 1, -1, CompareMethod.Text)
        '    'ensure item is available
        '    wdDoc.SaveAs2(FileName:=strR, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocumentMacroEnabled, CompatibilityMode:=Microsoft.Office.Interop.Word.WdCompatibilityMode.wdCurrent)
        'Else
        '    strExt = ".docx"
        '    strR = Replace(strpathT, ".xml", strExt, 1, -1, CompareMethod.Text)
        '    'ensure item is available
        '    wdDoc.SaveAs2(FileName:=strR, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode:=Microsoft.Office.Interop.Word.WdCompatibilityMode.wdCurrent)
        'End If

        ''close Word
        'Try
        '    wdDoc.Close(False)
        '    wd.Quit()
        'Catch ex As Exception
        '    var1 = ex.Message
        '    var1 = var1
        'End Try

        'Createxml = strR


    End Function

    Private Sub cbxFR1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxFR1.SelectedIndexChanged

        If boolFormLoad Or boolHold Then
            Exit Sub
        End If

    End Sub

    Private Sub cbxFR2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxFR2.SelectedIndexChanged

        If boolFormLoad Or boolHold Then
            Exit Sub
        End If

    End Sub

    Private Sub cmdOpenPDF_Click(sender As Object, e As EventArgs) Handles cmdOpenPDF.Click

        Call DoPDF()

end2:

    End Sub

    Sub DoPDF()

        Dim strM As String
        Dim var1

        'If Me.ovDC1.IsOpened Then
        If AllowWord() Then
        Else
            strM = "A document must be loaded."
            MsgBox(strM, vbInformation, "Invalid action...")
            GoTo end1
        End If


        If boolTemplate Then 'Report Template
            If BOOLALLOWREPORTTEMPLATEPDF Then
            Else
                strM = "User does not have permissions to create a PDF file."
                MsgBox(strM, vbInformation, "Invalid action...")
                GoTo end1
            End If
        Else
            If BOOLALLOWPDFREPORT Or BOOLFORCEFINALREPORTPDF Then
            Else
                strM = "User does not have permissions to create a PDF file."
                MsgBox(strM, vbInformation, "Invalid action...")
                GoTo end1
            End If
        End If

        If GoToWordPrep(True) Then
        Else
            GoTo end2
        End If

        If ShowOptionsForm("PDF") Then
        Else
            Exit Sub
        End If

        If ESigPrompt() Then
        Else
            GoTo end2
        End If

        Me.lblProgress1.Text = "Opening document as PDF..."
        Me.panProgress1.Visible = True
        Me.panProgress1.Refresh()

        Dim boolDo As Boolean = False
        Try

            If UBound(arrRanges1) = 1000 Then
            Else

                'need to reset ranges
                Call RestoreProtectedRanges(Me.ovDC1, gid1)
                boolDo = True

            End If

        Catch ex As Exception
            var1 = ex.Message
            var1 = var1

        End Try

        Call DoWrite()

        Call GoToWord(True)

        Try

            If boolDo Then

                'need to store ranges
                Call StoreProtectedRanges(Me.ovDC1, gid1)

            End If

        Catch ex As Exception
            var1 = ex.Message
            var1 = var1

        End Try

        Call DoReadOnly(Me.ovDC1)

end1:

        Me.lblProgress1.Text = ""
        Me.panProgress1.Visible = False
        Me.panProgress1.Refresh()

end2:

    End Sub

    Private Sub cmdWord_Click(sender As Object, e As EventArgs) Handles cmdWord.Click

        Dim strM As String
        Dim str1 As String
        Dim var1

        'If Me.ovDC1.IsOpened Then
        If AllowWord() Then
        Else
            strM = "A document must be loaded."
            MsgBox(strM, vbInformation, "Invalid action...")
            GoTo end1
        End If

        If boolTemplate Then 'Report Template
            If BOOLALLOWREPORTTEMPLATEWORD Then
            Else
                strM = "User does not have permissions to open this document in Word."
                MsgBox(strM, vbInformation, "Invalid action...")
                GoTo end1
            End If
        Else
            If BOOLFORCEFINALREPORTPDF Then

                str1 = "User permissions are set to force document as PDF."
                str1 = str1 & ChrW(10) & ChrW(10) & "Therefore, user does not have permissions to open this document in Word."
                strM = str1
                MsgBox(strM, vbInformation, "Invalid action...")
                GoTo end1
            Else
                If BOOLALLOWFINALREPORTWORD Then
                Else
                    strM = "User does not have permissions to open this document in Word."
                    MsgBox(strM, vbInformation, "Invalid action...")
                    GoTo end1
                End If
            End If

        End If

        If GoToWordPrep(False) Then
        Else
            GoTo end2
        End If

        If ShowOptionsForm("Word") Then
        Else
            Exit Sub
        End If

        If ESigPrompt() Then
        Else
            GoTo end2
        End If

        Me.lblProgress1.Text = "Opening document in Word..."
        Me.panProgress1.Visible = True
        Me.panProgress1.Refresh()

        Dim boolDo As Boolean = False
        Try

            If boolReadOnlyOp Then
                Call DoReadOnly(Me.ovDC1)
            Else
                If UBound(arrRanges1) = 1000 Then
                Else

                    'need to reset ranges
                    Call RestoreProtectedRanges(Me.ovDC1, gid1)
                    boolDo = True

                End If
            End If

        Catch ex As Exception
            var1 = ex.Message
            var1 = var1

        End Try

        If boolReadOnlyOp And boolNoneOp Then
        Else
            Call DoWrite()
        End If

        Call GoToWord(False)

        Try

            If boolDo Then

                'need to store ranges
                If boolReadOnlyOp Then
                Else
                    Call StoreProtectedRanges(Me.ovDC1, gid1)
                End If

            End If

        Catch ex As Exception
            var1 = ex.Message
            var1 = var1

        End Try

        Call DoReadOnly(Me.ovDC1)

        'Call ClearTemp()

end1:

        Me.lblProgress1.Text = ""
        Me.panProgress1.Visible = False
        Me.panProgress1.Refresh()

end2:

    End Sub

    Function GoToWordPrep(boolPDF As Boolean) As Boolean

        GoToWordPrep = False

        Dim strM As String
        Dim var1

        If Me.cmdSave.Visible And gGoToWord = False Then
            'strM = "Are you sure you wish to open this document in Microsoft" & ChrW(174) & " Word?" & ChrW(10) & ChrW(10)
            'strM = strM & "If any changes have not been saved, they will be lost."

            If boolTemplate Then 'this is Report Template

                If boolPDF Then
                    strM = "The Loaded Document will be opened as a .pdf file." & ChrW(10) & ChrW(10)
                    strM = strM & "This StudyDoc window will remain open."
                Else
                    strM = "The Loaded Document will be opened in Microsoft" & ChrW(174) & " Word." & ChrW(10) & ChrW(10)
                    strM = strM & "This StudyDoc window will remain open."
                End If

                var1 = MsgBox(strM, MsgBoxStyle.OkCancel, "Are you sure...")
                If var1 = 1 Then
                Else
                    GoTo end1
                End If

            Else
                If BOOLFORCEFINALREPORTPDF And boolFromReportHistory = False Then

                Else
                    If boolPDF Then
                        strM = "The Loaded Document will be opened as a .pdf file." & ChrW(10) & ChrW(10)
                        strM = strM & "This StudyDoc window will remain open."
                    Else
                        strM = "The Loaded Document will be opened in Microsoft" & ChrW(174) & " Word." & ChrW(10) & ChrW(10)
                        strM = strM & "This StudyDoc window will remain open."
                    End If

                    var1 = MsgBox(strM, MsgBoxStyle.OkCancel, "Are you sure...")
                    If var1 = 1 Then
                    Else
                        GoTo end1
                    End If
                End If
            End If

        End If

        GoToWordPrep = True

end1:

    End Function

    Sub GoToWord(boolPDF As Boolean)

        'Try
        '    Me.ovDC1.UnProtectDoc()
        'Catch ex As Exception

        'End Try

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

        Dim var1
        Dim strM As String
        Dim strExt As String

        Dim str1 As String
        Dim str2 As String


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


        Dim strPathO As String

        boolM = False
        Try

            'save the current document
            Me.ovDC1.Save()

            wdDoc = Me.ovDC1.ActiveDocument
            strPathO = wdDoc.FullName

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

        'do not close anymore

        'Pause(0.25)

        'Me.Visible = False


        'Try
        '    'Me.afrWord.Close()
        '    Me.ovDC1.CloseDoc() 'v8
        'Catch ex As Exception

        'End Try


        'Try
        '    '20140228
        '    'added this because lots of winword processes open after closing frmWordStatement
        '    Me.ovDC1.Dispose()
        'Catch ex As Exception

        'End Try

        'Me.ov1.Close() 'v6

        'now close current ov1 and open original
        Try
            Me.ovDC1.CloseDoc(False)
            Pause(0.25)
        Catch ex As Exception
        End Try


        Me.ovDC1.OpenWord(strPathO)

        'now open new path
        Dim wd As New Microsoft.Office.Interop.Word.Application
        Dim doc As Microsoft.Office.Interop.Word.Document

        Try
            wd.Application.NormalTemplate.Saved = True
        Catch ex As Exception

        End Try

        Try
            If boolM Then
                wd.Documents.Open(FileName:=strR)
            Else
                wd.Documents.Open(FileName:=strR)
            End If

            doc = wd.ActiveDocument

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

        If boolNoneOp Then
        Else

            Try

                str1 = FooterText()

                Call DoDocLabel(wd, doc, str1)

                If boolReadOnlyOp Then
                    'can't call readonly, must protect in word
                    Dim strPW As String = RandomPswd()
                    doc.Protect(Password:=strPW, NoReset:=False, Type:=Microsoft.Office.Interop.Word.WdProtectionType.wdAllowOnlyReading, UseIRM:=False, EnforceStyleLock:=False)
                End If

            Catch ex As Exception
                var1 = ex.Message
                var1 = var1
            End Try

        End If

        If boolPDF Then
            Dim strP As String

            strP = CreatePDF(wd, strR)

            Try
                wd.Application.ShowWindowsInTaskbar = boolSTB
            Catch ex As Exception

            End Try

            'must pause here or Adobe has a spaz
            Pause(0.25)
            wd.Quit(False)

        Else
            Try
                wd.Application.ShowWindowsInTaskbar = boolSTB
            Catch ex As Exception
                var1 = var1 'debug
            End Try
            wd.Visible = True
            wd.ActiveWindow.WindowState = Office.Interop.Word.WdWindowState.wdWindowStateMaximize

            '20180816 LEE:
            'If document is locked, View is outline
            'needs to be wdPrintview
            Try
                wd.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
            Catch ex As Exception
                var1 = var1
            End Try

            Try
                wd.ActiveWindow.Activate()
            Catch ex As Exception

            End Try
            'Pause(2.5)

        End If

        If boolTemplate Then
        Else
            If BOOLFORCEFINALREPORTPDF And boolFromReportHistory = False Then
                Call ExitDoc()
            End If
        End If

        'do not close anymore
        'Me.Close()
        'Me.Dispose()


    End Sub

    Sub DoDocLabel(wd As Microsoft.Office.Interop.Word.Application, doc As Microsoft.Office.Interop.Word.Document, strM As String)

        Dim sec As Microsoft.Office.Interop.Word.Section
        Dim intSec As Int16
        Dim str1 As String
        Dim str2 As String = Me.lblProgress1.Text
        Dim str3 As String
        Dim var1

        Dim numSecs As Int16

        numSecs = doc.Sections.Count

        wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

        'wd.Visible = True

        Dim strDate As String
        Dim strTime As Date
        Dim dt As Date
        dt = Now
        strDate = Format(dt, "MMM dd, yyyy")
        strTime = Format(dt, "hh:mm:ss tt")
        'str1 = "DRAFT" & Chr(10) & strDate & Chr(10) & strTime
        str1 = "DRAFT" & Chr(10) & strDate & " " & strTime
        If Len(strM) = 0 Then
            strM = str1
        Else
            strM = str1 & ChrW(10) & strM
        End If


        Call InsertWatermark(wd, True, strM)

        Exit Sub

        Try
            With wd

                If boolNoneOp Then
                Else

                    Try
                        If wd.ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdNormalView Or wd.ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdOutlineView Then
                            wd.ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
                        End If
                    Catch ex As Exception
                        var1 = var1
                    End Try
                  

                    '.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageFooter

                    For Each sec In doc.Sections

                        intSec = sec.Index

                        str1 = "Entering Document Label for Section " & intSec & " of " & numSecs
                        str3 = str2 & ChrW(10) & ChrW(10) & str1
                        Me.lblProgress1.Text = str3
                        Me.lblProgress1.Refresh()

                        If intSec = 1 Then
                            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToSection, Which:=Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToFirst, Count:=intSec, Name:="")
                            .ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageFooter
                        End If

                        If intSec = 1 Or sec.Footers(1).LinkToPrevious = False Then

                            '.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageFooter

                            If boolTextOp Then
                                .Selection.WholeStory()
                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                                .Selection.TypeParagraph()
                                'decrease font size
                                .Selection.Font.Size = .Selection.Font.Size - 2
                                .Selection.TypeText(Text:=strM)
                            ElseIf boolWaterMarkOp Then
                                Call InsertWatermarkDC(wd, doc, sec, intSec, strM)
                            End If

                            var1 = var1 'debug



                        Else

                        End If

                        Try
                            doc.ActiveWindow.ActivePane.View.NextHeaderFooter()
                        Catch ex As Exception

                        End Try


                        'If boolTextOp Then

                        '    If intSec = 1 Or sec.Headers.Item(Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).LinkToPrevious = False Then
                        '        .ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageFooter
                        '        .Selection.WholeStory()
                        '        .Selection.MoveRight(Unit:=WdUnits.wdCharacter, Count:=1)
                        '        .Selection.TypeParagraph()
                        '        'decrease font size
                        '        .Selection.Font.Size = .Selection.Font.Size - 2
                        '        .Selection.TypeText(Text:=strM)
                        '    Else

                        '    End If

                        'ElseIf boolWaterMarkOp Then

                        '    If boolCenterOp Then

                        '        Call InsertWatermarkDC(wd, doc, sec, intSec, strM)

                        '    ElseIf boolFooterOp Then

                        '    End If

                        'End If

                    Next

                    str1 = "Entering Document Label for Section " & numSecs & " of " & numSecs
                    str3 = str2 & ChrW(10) & ChrW(10) & str1
                    Me.lblProgress1.Text = str3
                    Me.lblProgress1.Refresh()

                    Me.lblProgress1.Text = str2
                    Me.lblProgress1.Refresh()

                End If



            End With
        Catch ex As Exception

            var1 = ex.Message
            var1 = var1

        End Try


        Try
            wd.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument
        Catch ex As Exception

        End Try

        Try
            Clipboard.Clear()
        Catch ex As Exception

        End Try

        wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

    End Sub


    Sub InsertWatermarkDC(ByVal wd As Microsoft.Office.Interop.Word.Application, doc As Microsoft.Office.Interop.Word.Document, sec As Microsoft.Office.Interop.Word.Section, intSec As Int16, strM As String)

        Dim var1, var2
        Dim str1 As String

        wd.Visible = True

        Dim strDate As String
        Dim strTime As Date
        Dim dt As Date
        dt = Now
        strDate = Format(dt, "MMM dd, yyyy")
        strTime = Format(dt, "hh:mm:ss tt")
        str1 = "DRAFT" & ChrW(10) & strDate & ChrW(10) & strTime
        strM = str1 & ChrW(10) & strM
        Try
            Call InsertWatermark(wd, True, strM)
        Catch ex As Exception
            var1 = var1
        End Try


    End Sub

    Function AllowWord() As Boolean

        AllowWord = False

        If Me.ovDC1.IsOpened Then
            AllowWord = True
        End If

end1:

    End Function


    Private Sub cmdEdit_Click(sender As Object, e As EventArgs) Handles cmdEdit.Click

        Dim strM As String

        If Me.rbCompare.Checked And Me.rbLoadCompare.Checked Then
            strM = "Document editing not available when Word Compare is active."
            MsgBox(strM, vbInformation, "Invalid action...")
            GoTo end1
        End If

        Dim boolDo As Boolean = DefaultFormats("Edit")

        If boolDo Then

            Call DoEdit("Edit")

        End If

end1:


    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click

        'load original
        'Call CompareFinalReports(False, True)

        Me.lblProgress1.Text = "Canceling edits..."
        Me.lblProgress1.Refresh()
        Me.panProgress1.Visible = True
        Me.panProgress1.Refresh()

        Call ResetLoad()

        Call DoEdit("Cancel")
        Call DefaultFormats("Cancel")

        Me.panProgress1.Visible = False
        Me.panProgress1.Refresh()
        Me.lblProgress1.Text = ""

    End Sub

    Private Sub cmdSave_Click(sender As Object, e As EventArgs) Handles cmdSave.Click

        Dim intR As Short
        Dim strM As String
        Dim str1 As String
        Dim str2 As String
        Dim intRow As Int16
        Dim id As Int64

        If AllowWord() Then
        Else
            strM = "A document must be loaded."
            MsgBox(strM, vbInformation, "Invalid action...")
            GoTo end1
        End If

        'If BOOLEDITFINALREPORT Then
        'Else
        '    strM = "User does not have permissions to Edit reports."
        '    strM = strM & ChrW(10) & "Therefore, user may not execute a Save action."
        '    MsgBox(strM, vbInformation, "Invalid action...")
        '    GoTo end1
        'End If


        'If Me.rbFinalReport.Checked Then
        '    str1 = "Final Report"
        '    strDocType = "Final"
        'Else
        '    str1 = "Section/Draft"
        '    strDocType = "Section"
        'End If

        'strM = "You have chosen to save this document as a '" & str1 & "'"
        'strM = strM & ChrW(10) & ChrW(10) & "If this is correct, click OK."
        'intR = MsgBox(strM, vbOKCancel, "Continue?")
        'If intR = 1 Then
        'Else
        '    GoTo end1
        'End If

        Dim frm As New frmSaveFinalDraft
        frm.ShowDialog()
        Dim boolCancel As Boolean = frm.boolCancel
        Dim boolFinal As Boolean = frm.rbFinal.Checked
        frm.Dispose()
        If boolCancel Then
            GoTo end1
        End If
        Me.rbFinalReport.Checked = boolFinal
        Me.rbSection.Checked = Not (boolFinal)

        If Me.rbFinalReport.Checked Then
            str1 = "Final Report"
            strDocType = "Final"
        Else
            str1 = "Section/Draft"
            strDocType = "Section"
        End If

        Dim strDescr As String = ""

        Dim var1
        Dim Count1 As Short

        Dim frm1 As New frmVersionDescr
        frm1.rtbD.Text = strDescr
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

        If ESigPrompt() Then
        Else
            GoTo end1
        End If

        Me.lblProgress1.Text = "Saving document..."
        Me.panProgress1.Visible = True
        Me.panProgress1.Refresh()

        '20181110 LEE:
        'protected docs have a problem with the next line
        'try disabling it
        'Call DoWrite()

        If boolHasBeenSaved Then
        Else
            'must restore ranges before saving
            'at this point, id = 0
            Call RestoreProtectedRanges(Me.ovDC1, 0)

            boolHasBeenSaved = True
        End If

        Call DoSave(strDescr)

        'make sure first row in dgvVersions
        Dim dgv As DataGridView
        For Count1 = 1 To 2
            Select Case Count1
                Case 1
                    dgv = Me.dgvFinalReports
                Case 2
                    dgv = Me.dgvSections
            End Select

            Try
                dgv.Rows(0).Selected = True
            Catch ex As Exception
                var1 = var1
            End Try

            Try
                dgv.CurrentCell = dgv.Item(GetVisibleCol(dgv), 0)
            Catch ex As Exception
                var1 = var1
            End Try

        Next

        If InStr(1, strDocType, "Final", CompareMethod.Text) > 0 Then
            id = Me.dgvFinalReports("ID_TBLFINALREPORT", 0).Value
        Else
            id = Me.dgvSections("ID_TBLFINALREPORT", 0).Value
        End If

        gid1 = id

        Me.gbLoad.Visible = True

        Me.rbLoad.Checked = True

        ''reset to load document
        'If Len(gDoc) = 0 Then
        '    gDoc = ""
        'Else
        '    Call ResetLoad()
        'End If

        'pesky
        Try
            Me.ovDC1.ActiveDocument.ActiveWindow.View.SplitSpecial = Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneNone
        Catch ex As Exception
            'MsgBox("ov.ActiveDocument.ActiveWindow.View.SplitSpecial" & ChrW(10) & ex.Message)
            var1 = var1
        End Try

        Try
            Me.ovDC1.ActiveDocument.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
        Catch ex As Exception
            var1 = var1
        End Try


        Try
            'Call DoReadOnly(Me.ovDC1)
            '20181110 LEE:
            'getting a Windows (officeviewer) hang error if document has embedded protected sections
            'try calling here instead of in subroutine
            'Me.ovDC1.ProtectDoc(EDOfficeLib.WdProtectType.wdAllowOnlyReading)
            'nope, still doesn't work
            'Turns out DoReadOnly isn't needed

        Catch ex As Exception
            var1 = var1
        End Try

        Me.lblInstructions01.Visible = False

        Dim strCW As String
        strCW = ReturnstrCW(gid1)
        Me.txtLoadedDocDescription.Text = strCW


end2:

        Call DoEdit("Save")

        Call DefaultFormats("Save")

end1:

        Me.panProgress1.Visible = False
        Me.panProgress1.Refresh()

    End Sub

    Sub SetlblInstructions()

        Dim str1 As String

        'set position
        Me.lblInstructions01.Top = Me.gbLoad.Top
        Me.lblInstructions01.Size = Me.gbLoad.Size

        str1 = "Document cannot be modified or printed until it has been Saved (instituting version control)." & ChrW(10)
        str1 = str1 & "Click 'Go Back' if the document is not to be Saved."

        Me.lblInstructions01.Text = str1



    End Sub

    Sub ResetLoad()

        Try

            Call CompareFinalReports(True, True, gid1)

        Catch ex As Exception

        End Try

        'Me.rbLoad.Checked = True

end1:


    End Sub

    Sub DoSave(strDescr As String)

        Dim strReportO As String

        Dim wdDoc As Microsoft.Office.Interop.Word.Document
        'Dim wdApp As Microsoft.Office.Interop.Word.Application
        Dim var1 As Object
        Dim strt

        Cursor.Current = Cursors.WaitCursor

        'Me.cmdCancel.Enabled = False
        'Me.cmdEditStatements.Enabled = True

        Dim oPath As String
        Dim oName As String
        Dim strUPath As String = "C:\LabIntegrity\StudyDoc\Temp" ' "C:\Labintegrity\StudyDoc\Temp"
        Dim strNPath As String
        Dim strPathT As String
        Dim Count1 As Int16

        Dim dgvFR As DataGridView = Me.dgvFinalReports
        Dim dgvSec As DataGridView = Me.dgvSections

        Try

            wdDoc = Me.ovDC1.ActiveDocument

            Try
                wdDoc.Revisions.AcceptAll()
            Catch ex As Exception
                var1 = var1
            End Try


            'strt = wdDoc.Application.Selection.Start

            'find path of wddoc
            oPath = wdDoc.Path
            oName = wdDoc.FullName

            'get new file
            strPathT = GetNewTempFile(True)
            'Me.lblSection.Text = strPathT

            'first save original doc
            wdDoc.Application.DisplayAlerts = False
            Try
                wdDoc.Save()
            Catch ex As Exception
                MsgBox("DoSave first save record" & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
            End Try


            Dim strExt As String
            Dim strR As String
            Dim bool2007 As Boolean
            Dim strPath As String
            Dim boolM As Boolean
            Dim ver As Short

            ver = CInt(wdDoc.Application.Version)
            If ver < 13 Then
                bool2007 = True
            Else
                bool2007 = False
            End If

            Try
                'now save as new file for fs to open
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

            Catch ex As Exception
                MsgBox("now save as new file for fs to open" & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
            End Try

            wdDoc.Application.DisplayAlerts = True

            'keep setting this

            Call SetWordStuff(wdDoc)

            'close the doc to release the document
            If Me.ovDC1.IsOpened Then
                Try
                    Me.ovDC1.CloseDoc(False)
                Catch ex As Exception

                End Try
            End If

            're-open original
            Try
                Me.ovDC1.OpenWord(oName)
            Catch ex As Exception
                var1 = var1
            End Try

            Cursor.Current = Cursors.WaitCursor

            Call UpdateDatabaseSave(strPathT, strDescr)


            Cursor.Current = Cursors.WaitCursor

        Catch ex As Exception

            Dim strM As String
            strM = "A problem occurred when attempting to open the Word file. Please try again." & ChrW(10) & ex.Message
            MsgBox(strM, MsgBoxStyle.Information, "Problem...")

        End Try


        dgvFR.AutoResizeRows()
        dgvSec.AutoResizeRows()


        Cursor.Current = Cursors.Default

        wdDoc = Nothing
        ' wdApp = Nothing

        'now make things visible
        'Me.gbSaveType.Visible = True

        Me.cmdCompareFinalReport.Visible = True
        Me.cmdCompareSection.Visible = True

        Call EnableControls(True)
end1:


        boolHold = False

        Cursor.Current = Cursors.Default

        Try
            If Me.rbFinalReport.Checked Then
                dgvFR.Rows(0).Selected = True
                Try
                    dgvFR.CurrentCell = dgvFR.Item(GetVisibleCol(dgvFR), 0)
                Catch ex As Exception
                    var1 = var1
                End Try
            Else
                dgvSec.Rows(0).Selected = True
                Try
                    dgvSec.CurrentCell = dgvSec.Item(GetVisibleCol(dgvSec), 0)
                Catch ex As Exception
                    var1 = var1
                End Try
            End If
        Catch ex As Exception

        End Try

        'enter label
        'create label
        Dim strCW As String
        Dim dgv As DataGridView
        strCW = "Loaded Document: "

        Try

            strCW = ReturnstrCW(gid1)
            Me.txtLoadedDocDescription.Text = strCW

        Catch ex As Exception

        End Try

        'clear compare
        'Me.txtComparedDocDescription.Text = ""


        'Me.Close()

    End Sub

    Sub EnableControls(bool As Boolean)

        Me.gbLoad.Enabled = bool
        Me.cmdCompareFinalReport.Enabled = bool
        Me.cmdCompareSection.Enabled = bool

    End Sub

    Sub UpdateDatabaseSave(ByVal strPathT As String, strD As String)

        'this routine will populate the table tblFinalReportWordDocs

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
        Dim intMaxFR As Int64

        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        Dim dt As Date
        Dim idHere As Int64
        Dim var1, var2
        Dim intVer As Int32 = 0
        Dim strReportType As String

        Dim intRow As Int32

        Dim dtNow As Date = Now

        'first add a new record to tblFinalReport

        'now get new maxid
        intMaxFR = GetMaxID("TBLFINALREPORT", 1, True) 'this returns incremented value and records incremented value

        'legend
        'ID_TBLFINALREPORT
        'ID_TBLSTUDIES
        'ID_TBLREPORTHISTORY
        'ID_TBLREPORTS
        'INTFINALREPORTVERSION
        'CHARDESCRIPTION
        'CHARCOMMENTS
        'CHARREPORTTYPE
        'BOOLLOCKED
        'ID_TBLPERSONNEL
        'ID_TBLUSERACCOUNTS
        'CHARUSERID
        'UPSIZE_TS

        If id_tblReports = 0 Then
            intRow = frmH.dgvReports.CurrentRow.Index
            id_tblReports = frmH.dgvReports("ID_TBLREPORTS", intRow).Value
        End If

        If id_tblReportHistory = 0 Then
            'may have to find this later
        End If

        Dim tblFR As System.Data.DataTable = tblFinalReport
        Dim nr As DataRow = tblFR.NewRow

        nr.BeginEdit()

        gCHARREPORTGENERATEDSTATUS = Me.txtDescr.Text

        nr("ID_TBLFINALREPORT") = intMaxFR
        nr("ID_TBLSTUDIES") = id_tblStudies
        nr("ID_TBLREPORTHISTORY") = id_tblReportHistory
        nr("ID_TBLREPORTS") = id_tblReports
        If Me.rbFinalReport.Checked Then
            'get version
            intVer = GetWordVersion(0, False)
            strReportType = "Final Report"
        Else
            intVer = 0
            strReportType = "Section"
        End If
        nr("INTFINALREPORTVERSION") = intVer + 1
        nr("CHARDESCRIPTION") = gCHARREPORTGENERATEDSTATUS
        nr("CHARCOMMENTS") = strD
        nr("CHARREPORTTYPE") = strReportType
        nr("BOOLLOCKED") = 0
        nr("ID_TBLPERSONNEL") = id_tblPersonnel
        nr("ID_TBLUSERACCOUNTS") = id_tblUserAccounts
        nr("CHARUSERID") = gUserID
        nr("UPSIZE_TS") = dtNow

        If Len(tPswd) = 0 Then
        Else
            nr("CHARPASSWORD") = tPswd 'encrypted
        End If

        nr.EndEdit()
        tblFR.Rows.Add(nr)

        'update frmh
        If StrComp(strReportType, "Final Report", CompareMethod.Text) = 0 Then
            str1 = LTextDateFormat & " HH:mm:ss tt"
            Dim strDt As String = Format(dtNow, str1)
            frmH.lblFinalReportLockedDate.Text = strDt
        End If

        If gboolAuditTrail Then

            Call FillAuditTrailTemp(tblFinalReport)
            'record tblaudittrailtemp
            Call RecordAuditTrail(True, dtNow)

        End If

        'update database
        If boolGuWuOracle Then
            'Try
            '    ta_TBLFINALREPORT.Update(TBLFINALREPORT)
            'Catch ex As DBConcurrencyException
            '    'ds2005.TBLSAMPLERECEIPT.Merge('ds2005.TBLSAMPLERECEIPT, True)
            'End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_TBLFINALREPORTAcc.Update(tblFinalReport)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLSAMPLERECEIPT.Merge('ds2005Acc.TBLSAMPLERECEIPT, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_TBLFINALREPORTSQLServer.Update(tblFinalReport)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLSAMPLERECEIPT.Merge('ds2005Acc.TBLSAMPLERECEIPT, True)
            End Try
        End If


        Dim strT As String

        Dim strPathTT As String

        Dim constr As String
        If boolGuWuAccess Then
            constr = constrIni
        ElseIf boolGuWuSQLServer Then
            constr = constrIni ' "Provider=SQLOLEDB;" & constrIni
        ElseIf boolGuWuOracle Then
            constr = constrIniGuWuODBC
        End If

        Dim myConnection As OleDb.OleDbConnection
        Dim myCommand As New OleDb.OleDbCommand

        myConnection = New OleDb.OleDbConnection(constr)
        myConnection.Open()
        myCommand.Connection = myConnection
        myCommand.CommandType = CommandType.Text


        'do tblFinalReportWordDoc

        dtbl1 = tblFinalReportWordDocs
        strT = "tblFinalReportWorddocs"
        dt = dtNow ' Now

        Dim boolE As Boolean

        Count1 = 0

        Dim fs As FileStream

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

        'legend
        'ID_TBLFINALREPORTWORDDOCS
        'ID_TBLFINALREPORT'intMaxFR
        'ID_TBLSTUDIES
        'INTFINALREPORTVERSION'intver
        'CHARREPORTTYPE'strReportType
        'CHARXML
        'UPSIZE_TS


        'Write in blocks of 2000 characters.
        'NDL Note: used "charCount - 1" as that is what was used previously.  
        '          "chars.Length" yields a different (slightly smaller) number.
        Dim intLastSegment As Int32 = Math.Truncate((charCount - 1) / 2000)

        'now get new maxid
        intMax = GetMaxID(strT, intLastSegment, True)

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
            strSQL = "INSERT INTO " & strT & "  VALUES (" & intMax & "," & intMaxFR & "," & id_tblStudies & "," & intVer & ",'" & strReportType & "','" & strW & "'," & ReturnDate(dtNow) & ");"

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

        Call PutMaxID(strT, intMax)


        myCommand.Connection = Nothing
        If myConnection.State = ADODB.ObjectStateEnum.adStateOpen Then
            myConnection.Close()
        End If
        myConnection = Nothing


end1:


    End Sub

    Private Sub dgvSections_MouseHover(sender As Object, e As EventArgs) Handles dgvSections.MouseHover

        'Cursor.Current = Cursors.Default

    End Sub

    Private Sub dgvFinalReports_MouseHover(sender As Object, e As EventArgs) Handles dgvFinalReports.MouseHover

        'Cursor.Current = Cursors.Default

    End Sub

    Private Sub cmdCompareFinalReport_Click(sender As Object, e As EventArgs) Handles cmdCompareFinalReport.Click

        Dim dgv As DataGridView = Me.dgvFinalReports
        If dgv.RowCount = 0 Then
            GoTo end1
        End If

        Dim intRow As Int16
        intRow = dgv.CurrentRow.Index


        Dim id As Int64
        id = dgv("ID_TBLFINALREPORT", intRow).Value

        If Me.rbLoad.Checked Then
            Dim strDesc As String
            strDesc = dgv("CHARDESCRIPTION", intRow).Value
            Me.txtDescr.Text = strDesc
            gid1 = id
        Else
            gid2 = id
        End If

        boolHasBeenClicked = True
        strDocType = "Final"
        Call CompareFinalReports(False, False, id)

        'Call DefaultFormats("Cancel")

end1:

    End Sub

    Private Sub cmdCompareSection_Click(sender As Object, e As EventArgs) Handles cmdCompareSection.Click

        Dim dgv As DataGridView = Me.dgvSections
        If dgv.RowCount = 0 Then
            GoTo end1
        End If

        Dim intRow As Int16
        intRow = dgv.CurrentRow.Index



        Dim id As Int64
        id = dgv("ID_TBLFINALREPORT", intRow).Value

        If Me.rbLoad.Checked Then
            Dim strDesc As String
            strDesc = dgv("CHARDESCRIPTION", intRow).Value
            Me.txtDescr.Text = strDesc
            gid1 = id
        Else
            gid2 = id
        End If

        boolHasBeenClicked = True
        strDocType = "Section"
        Call CompareFinalReports(True, False, id)
        'Call DefaultFormats("Cancel")
        intRow = intRow 'debug
end1:

    End Sub

    Sub CompareFinalReports(boolSection As Boolean, boolFromCancel As Boolean, id As Int64)

        Dim str1 As String
        Dim str2 As String
        Dim strCW1 As String
        Dim intRow As Int16
        Dim boolLoad As Boolean
        Dim var1
        'Dim id As Int64
        Dim ov As AxEDOfficeLib.AxEDOffice
        Dim strM As String

        '20190115 LEE:
        Dim boolWC As Boolean 'If Word Compare is chosen
        boolWC = Me.rbCompare.Checked

        Dim boolEnd As Boolean = False

        Call PlaceProgress()

        If boolFromCancel Then

            id = gid1
            ov = Me.ovDC1
            Call LoadAFR(True, boolSection, boolFromCancel)

        Else

            If boolSection Then
                str1 = Me.cmdCompareSection.Text
            Else
                str1 = Me.cmdCompareFinalReport.Text
            End If

            Dim boolLoaded As Boolean = Me.rbLoad.Checked

            'If InStr(1, str1, "Compare", CompareMethod.Text) > 0 And boolLoaded Then
            If boolLoaded Then
                boolLoad = True
            Else
                boolLoad = False
            End If

            Dim strF As String
            strF = "ID_TBLFINALREPORT = " & id
            Dim rows() As DataRow = tblFinalReport.Select(strF)

            If boolLoad Then
                gid1 = rows(0).Item("ID_TBLFINALREPORT")
                id = gid1
                Call LoadAFR(True, boolSection, boolFromCancel)
                ov = Me.ovDC1
            Else

                'check to see if document has been loaded
                If Me.ovDC1.IsOpened Then
                Else
                    strM = "A document must first be loaded."
                    MsgBox(strM, vbInformation, "Invalid action...")
                    boolEnd = True
                    GoTo end1
                End If

                gid2 = rows(0).Item("ID_TBLFINALREPORT")
                id = gid2
                Call LoadCompareFinalReport(boolSection)
                'pesky
                Call SetCompareOV(2, False)
                ov = Me.ovDC2
            End If

        End If

        'pesky
        'Call DoReviewPane()

        'If Me.rbLoad.Checked Then
        '    Call StoreProtectedRanges(Me.ovDC1, id)
        'End If
        'If boolFromCancel Then
        'Else
        '    Call StoreProtectedRanges(ov, id)
        'End If
        'Call StoreProtectedRanges(ov, id)

        Me.ovDC1.Visible = True

        If boolLoad Or boolFromCancel Then

            Call StoreProtectedRanges(ov, id)

            Try
                Me.ovDC.CloseDoc(False)
            Catch ex As Exception
                var1 = var1
            End Try

            'Try
            '    Me.ovDC2.CloseDoc(False)
            'Catch ex As Exception
            '    var1 = var1
            'End Try

            'Me.txtComparedDocDescription.Text = ""

            str1 = "Loaded Doc:"
            str2 = "Compared Doc:"
            strCW1 = "Loaded Document:"

            Me.lblNewDoc.Text = strCW1
            Me.lblLoadedDoc.Text = str1
            Me.lblComparedDoc.Text = str2

            Call ClearTemp()

        End If

        Call SpecialFormats()

        Call HilightSections()

end1:

        Me.panProgress1.Visible = False
        Me.panProgress1.Refresh()
        Me.lblProgress1.Text = ""

        If boolWC And boolEnd = False And Len(Me.txtComparedDocDescription.Text) > 0 Then '20190115 LEE:
            'call WordCompare
            Dim boolA As Boolean
            boolA = DoWordCompare()
        End If

end2:


    End Sub


    Sub LoadCompareFinalReport(boolSection As Boolean)

        Dim strF As String
        Dim intV1 As Int64
        Dim intV2 As Int64
        Dim strPath1 As String
        Dim strPath2 As String
        Dim strPathC As String
        Dim id As Int64
        Dim id1 As Int64 = gid1
        Dim id2 As Int64 = gid2

        Cursor.Current = Cursors.WaitCursor

        Call OpenFinalReportWordDocs(id1, 0, "", id2, 0, "")

        Call LoadAFR(True, boolSection, False)

        'Call ClearTemp()

        Cursor.Current = Cursors.Default
        'Me.cmdCompare.Enabled = False

    End Sub

    Private Sub rbSbyS_CheckedChanged(sender As Object, e As EventArgs) Handles rbSbyS.CheckedChanged

        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strCW1 As String
        Dim strCW2 As String
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short

        Dim var1

        Dim strM As String

        '20190115 LEE:
        'Why do any of this????
        'Instead, check for content and call DoWordCompare if necessary

        If Me.rbSbyS.Checked Then
            'do nothing
        Else
            str1 = Me.txtLoadedDocDescription.Text
            str2 = Me.txtComparedDocDescription.Text

            If Len(str1) = 0 Or Len(str2) = 0 Then
            Else
                Dim boolA As Boolean
                boolA = DoWordCompare()
            End If

        End If

        GoTo end1

        'strM = "Please note that,if a Side-By-Side Compared Document is already shown, the Compared document must be reloaded in order to show the Word Compare display."
        'Try
        '    If Me.rbCompare.Checked And Me.ovDC2.Visible And Me.ovDC2.IsOpened Then
        '        MsgBox(strM, vbInformation, "Note...")
        '    End If
        'Catch ex As Exception
        '    var1 = ex.Message
        '    var1 = var1
        'End Try

        Cursor.Current = Cursors.WaitCursor

        If Me.rbSbyS.Checked Then
            Me.chkRPane.Enabled = False
            str1 = "Loaded Doc:"
            str2 = "Compared Doc:"
            strCW1 = "Loaded Document:"
            Me.lblProgress1.Text = "Preparing Side-By-Side Compare..."
        Else
            Me.chkRPane.Enabled = True
            str1 = "Orginal Doc:"
            str2 = "Revised Doc:"
            strCW1 = "Original Document:"
            Me.lblProgress1.Text = "Preparing Word Compare..."
        End If
        Me.panProgress1.Visible = True
        Me.panProgress1.Refresh()

        Me.lblNewDoc.Text = strCW1
        Me.lblLoadedDoc.Text = str1
        Me.lblComparedDoc.Text = str2


        Me.ovDC1.Visible = False

        Call EnterLabels()

        Call SizePanes(True)

        Me.ovDC1.Visible = True


        str1 = Me.txtComparedDocDescription.Text

        If Len(str1) = 0 Then
            Try
                Me.ovDC2.CloseDoc(False)
            Catch ex As Exception

            End Try
            Me.ovDC2.Visible = False
        Else
            Me.ovDC2.Visible = True
        End If

        If Me.rbCompare.Checked Then
            If Len(str1) = 0 Then
            Else
                If Me.rbSbyS.Checked Then
                    'find ID and load DC2
                    int1 = InStr(1, str1, "ID:", CompareMethod.Text)
                    int2 = int1 + 4
                    If int1 = 0 Then
                    Else
                        'find next space
                        int3 = InStr(int2, str1, " ", CompareMethod.Text)
                        var1 = Mid(str1, int2, int3 - int2)
                        'MsgBox(var1.ToString)
                    End If
                End If
            End If
        End If

end1:

        Cursor.Current = Cursors.Default

        Me.panProgress1.Visible = False
        Me.panProgress1.Refresh()
        Me.lblProgress1.Text = ""


    End Sub

    Private Sub rbLoad_CheckedChanged(sender As Object, e As EventArgs) Handles rbLoad.CheckedChanged

        If boolFromClearCompare Then
        Else
            If Me.rbLoad.Checked Then
                Me.lblProgress1.Text = "Preparing Load..."
            Else
                Me.lblProgress1.Text = "Preparing Compare..."
            End If

            Me.panProgress1.Visible = True
            Me.panProgress1.Refresh()
        End If

        Call SizePanes(True)

        Me.ovDC1.Visible = True
        Me.ovDC2.Visible = True

        Call ShowButtons()

        'pesky
        Try
            If Me.ovDC2.IsOpened Then
                Me.cmdClearCompare.Enabled = True
                Me.cmdClearCompare.BackColor = System.Drawing.Color.Gainsboro
            End If
        Catch ex As Exception

        End Try

        If boolFromClearCompare Then
        Else
            Me.panProgress1.Visible = False
            Me.panProgress1.Refresh()
        End If


    End Sub


    Sub ShowButtons()

        Dim boolShow As Boolean = False
        Dim str1 As String = "Load-->"

        If Me.rbLoadCompare.Checked Then
            boolShow = True
            str1 = "Compare-->"
        End If

        'Me.cmdCompareFinalReport.Visible = boolShow
        'Me.cmdCompareSection.Visible = boolShow

        If Me.rbLoadCompare.Checked Then
            Me.gbCompare.Enabled = True
        Else
            Me.gbCompare.Enabled = False
        End If

        Me.cmdCompareFinalReport.Text = str1
        Me.cmdCompareSection.Text = str1

    End Sub

    Sub SizePanes(boolInvis As Boolean)

        Dim ns As Single = 100

        'if a compare doc is loaded, don't collapse it

        Dim str1 As String = Me.txtComparedDocDescription.Text

        Dim boolCol2 As Boolean

        If Len(str1) = 0 Then
            boolCol2 = True
        Else
            boolCol2 = False
        End If

        If Me.rbSbyS.Checked Then
            If Me.rbLoadCompare.Checked Then
                Me.sc1.Panel2Collapsed = False
            Else
                If boolCol2 Then
                    Me.sc1.Panel2Collapsed = True
                Else
                    Me.sc1.Panel2Collapsed = False
                End If
            End If
        Else
            Me.sc1.Panel2Collapsed = True
        End If

        Call SetCompareOV(3, boolInvis)

        'for some reason, ovdc1 right anchor is messed up
        Dim a, b, c, d

        a = Me.sc1.Panel1.Height
        b = Me.sc1.Panel1.Width
        c = Me.ovDC1.Left
        d = b - c - c

        Me.ovDC1.Width = d


    End Sub

    Sub DoReviewPane()

        '20190115 LEE:
        'Deprecated
        Exit Sub

        If boolTemplate Then
        Else
            If rbSbyS.Checked Or Me.rbLoad.Checked Then
                Exit Sub
            End If
        End If

        Dim var1
        Dim Count1 As Short
        Dim ov As AxEDOfficeLib.AxEDOffice
        Dim strCFR As String

        If boolTemplate Then
            ov = Me.ovDC
        Else
            ov = Me.ovDC1
        End If

        If boolTemplate Then
            If Me.chkRPane1.Checked Then

                Try
                    Try
                        ov.ActiveDocument.ActiveWindow.View.SplitSpecial = Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneRevisionsVert
                    Catch ex As Exception
                        'MsgBox("ov.ActiveDocument.ActiveWindow.View.SplitSpecial" & ChrW(10) & ex.Message)
                        var1 = var1
                    End Try
                Catch ex As Exception

                End Try

                Try
                    ov.ActiveDocument.ActiveWindow.ShowSourceDocuments = Microsoft.Office.Interop.Word.WdShowSourceDocuments.wdShowSourceDocumentsBoth
                Catch ex As Exception
                    'MsgBox("ov.ActiveDocument.ActiveWindow.ShowSourceDocuments" & ChrW(10) & ex.Message)
                    var1 = var1
                End Try

                Try
                    With ov.ActiveDocument.ActiveWindow.Document.Frameset.ChildFramesetItem(2)
                        .WidthType = Microsoft.Office.Interop.Word.WdFramesetSizeType.wdFramesetSizeTypePercent
                        .Width = 85
                    End With
                Catch ex As Exception
                    var1 = var1
                End Try

                Try
                    With ov.ActiveDocument.ActiveWindow.Document.Frameset.ChildFramesetItem(2).ChildFramesetItem(1)
                        .WidthType = Microsoft.Office.Interop.Word.WdFramesetSizeType.wdFramesetSizeTypePercent
                        .Width = 42.5
                    End With
                Catch ex As Exception
                    var1 = var1
                End Try
            Else

                Try
                    ov.ActiveDocument.ActiveWindow.View.SplitSpecial = Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneNone
                Catch ex As Exception
                    'MsgBox("ov.ActiveDocument.ActiveWindow.View.SplitSpecial" & ChrW(10) & ex.Message)
                    var1 = var1
                End Try

            End If
        Else
            If Me.chkRPane.Checked Then

                Try
                    Try
                        ov.ActiveDocument.ActiveWindow.View.SplitSpecial = Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneRevisionsVert
                    Catch ex As Exception
                        'MsgBox("ov.ActiveDocument.ActiveWindow.View.SplitSpecial" & ChrW(10) & ex.Message)
                        var1 = var1
                    End Try
                Catch ex As Exception

                End Try

                Try
                    ov.ActiveDocument.ActiveWindow.ShowSourceDocuments = Microsoft.Office.Interop.Word.WdShowSourceDocuments.wdShowSourceDocumentsBoth
                Catch ex As Exception
                    'MsgBox("ov.ActiveDocument.ActiveWindow.ShowSourceDocuments" & ChrW(10) & ex.Message)
                    var1 = var1
                End Try

                Try
                    With ov.ActiveDocument.ActiveWindow.Document.Frameset.ChildFramesetItem(2)
                        .WidthType = Microsoft.Office.Interop.Word.WdFramesetSizeType.wdFramesetSizeTypePercent
                        .Width = 85
                    End With
                Catch ex As Exception
                    var1 = var1
                End Try

                Try
                    With ov.ActiveDocument.ActiveWindow.Document.Frameset.ChildFramesetItem(2).ChildFramesetItem(1)
                        .WidthType = Microsoft.Office.Interop.Word.WdFramesetSizeType.wdFramesetSizeTypePercent
                        .Width = 42.5
                    End With
                Catch ex As Exception
                    var1 = var1
                End Try
            Else

                Try
                    ov.ActiveDocument.ActiveWindow.View.SplitSpecial = Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneNone
                Catch ex As Exception
                    'MsgBox("ov.ActiveDocument.ActiveWindow.View.SplitSpecial" & ChrW(10) & ex.Message)
                    var1 = var1
                End Try

            End If
        End If

        Try
            ov.ActiveDocument.ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
        Catch ex As Exception
            var1 = var1
        End Try

    End Sub

    Sub SetCompareOV(intI As Short, boolInVis As Boolean)

        '20190116 LEE:
        'Word Compare mode no longer exists
        'Hmmm. This code isn't needed anymore

        Exit Sub

        Dim var1
        Dim Count1 As Short
        Dim ov As AxEDOfficeLib.AxEDOffice
        Dim strCFR As String

        If boolTemplate Then
            ov = Me.ovDC
        Else
            ov = Me.ovDC1
        End If

        If boolInVis Then
            ov.Visible = False
        End If

        Try

            '20190116 LEE:
            'New logic
            'Word Compare mode no longer exists

            Try
                ov.ActiveDocument.ActiveWindow.View.SplitSpecial = Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneNone
            Catch ex As Exception
                'msgbox("ov.ActiveDocument.ActiveWindow.View.SplitSpecial" & ChrW(10) & ex.Message)
                var1 = ex.Message
                var1 = var1
            End Try

            Try
                ov.ActiveDocument.ActiveWindow.ShowSourceDocuments = Microsoft.Office.Interop.Word.WdShowSourceDocuments.wdShowSourceDocumentsOriginal
            Catch ex As Exception
                'msgbox("ov.ActiveDocument.ActiveWindow.ShowSourceDocuments" & ChrW(10) & ex.Message)
                var1 = ex.Message
                var1 = var1
            End Try

            Try
                With ov.ActiveDocument.ActiveWindow.Document.Frameset.ChildFramesetItem(1)
                    .WidthType = Microsoft.Office.Interop.Word.WdFramesetSizeType.wdFramesetSizeTypeFixed
                    .Width = 1
                End With
            Catch ex As Exception
                var1 = ex.Message
                var1 = var1
                'msgbox("ov.ActiveDocument.ActiveWindow.Document.Frameset" & ChrW(10) & ex.Message)
            End Try


        Catch ex As Exception

        End Try

        For Count1 = 1 To intI
            'ovdc1 may be in Word Compare mode
            '20190116 LEE: Word Compare mode no longer exists

            If boolTemplate Then

                Try
                    ov.ActiveDocument.ActiveWindow.View.SplitSpecial = Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneRevisionsVert
                Catch ex As Exception
                    'msgbox("ov.ActiveDocument.ActiveWindow.View.SplitSpecial" & ChrW(10) & ex.Message)
                    var1 = ex.Message
                    var1 = var1
                End Try

                Try
                    ov.ActiveDocument.ActiveWindow.ShowSourceDocuments = Microsoft.Office.Interop.Word.WdShowSourceDocuments.wdShowSourceDocumentsBoth
                Catch ex As Exception
                    'msgbox("ov.ActiveDocument.ActiveWindow.ShowSourceDocuments" & ChrW(10) & ex.Message)
                    var1 = ex.Message
                    var1 = var1
                End Try

                If Count1 > 0 Then
                    Try
                        With ov.ActiveDocument.ActiveWindow.Document.Frameset.ChildFramesetItem(2)
                            .WidthType = Microsoft.Office.Interop.Word.WdFramesetSizeType.wdFramesetSizeTypePercent
                            .Width = 85
                        End With
                    Catch ex As Exception
                        var1 = ex.Message
                        var1 = var1
                    End Try

                    Try
                        With ov.ActiveDocument.ActiveWindow.Document.Frameset.ChildFramesetItem(2).ChildFramesetItem(1)
                            .WidthType = Microsoft.Office.Interop.Word.WdFramesetSizeType.wdFramesetSizeTypePercent
                            .Width = 42.5
                        End With
                    Catch ex As Exception
                        var1 = ex.Message
                        var1 = var1
                        'msgbox("ov.ActiveDocument.ActiveWindow.Document.Frameset" & ChrW(10) & ex.Message)
                    End Try
                End If
            Else

                If boolInVis Then
                    ov.Visible = False
                End If

                Try

                    '20190116 LEE:
                    'New logic
                    Try
                        ov.ActiveDocument.ActiveWindow.View.SplitSpecial = Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneNone
                    Catch ex As Exception
                        'msgbox("ov.ActiveDocument.ActiveWindow.View.SplitSpecial" & ChrW(10) & ex.Message)
                        var1 = ex.Message
                        var1 = var1
                    End Try

                    Try
                        ov.ActiveDocument.ActiveWindow.ShowSourceDocuments = Microsoft.Office.Interop.Word.WdShowSourceDocuments.wdShowSourceDocumentsOriginal
                    Catch ex As Exception
                        'msgbox("ov.ActiveDocument.ActiveWindow.ShowSourceDocuments" & ChrW(10) & ex.Message)
                        var1 = ex.Message
                        var1 = var1
                    End Try

                    Try
                        With ov.ActiveDocument.ActiveWindow.Document.Frameset.ChildFramesetItem(1)
                            .WidthType = Microsoft.Office.Interop.Word.WdFramesetSizeType.wdFramesetSizeTypeFixed
                            .Width = 1
                        End With
                    Catch ex As Exception
                        var1 = ex.Message
                        var1 = var1
                        'msgbox("ov.ActiveDocument.ActiveWindow.Document.Frameset" & ChrW(10) & ex.Message)
                    End Try

                    'If Me.rbSbyS.Checked Then
                    '    Try
                    '        ov.ActiveDocument.ActiveWindow.View.SplitSpecial = Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneNone
                    '    Catch ex As Exception
                    '        'msgbox("ov.ActiveDocument.ActiveWindow.View.SplitSpecial" & ChrW(10) & ex.Message)
                    '        var1 = ex.Message
                    '        var1 = var1
                    '    End Try

                    '    Try
                    '        ov.ActiveDocument.ActiveWindow.ShowSourceDocuments = Microsoft.Office.Interop.Word.WdShowSourceDocuments.wdShowSourceDocumentsOriginal
                    '    Catch ex As Exception
                    '        'msgbox("ov.ActiveDocument.ActiveWindow.ShowSourceDocuments" & ChrW(10) & ex.Message)
                    '        var1 = ex.Message
                    '        var1 = var1
                    '    End Try

                    '    Try
                    '        With ov.ActiveDocument.ActiveWindow.Document.Frameset.ChildFramesetItem(1)
                    '            .WidthType = Microsoft.Office.Interop.Word.WdFramesetSizeType.wdFramesetSizeTypeFixed
                    '            .Width = 1
                    '        End With
                    '    Catch ex As Exception
                    '        var1 = ex.Message
                    '        var1 = var1
                    '        'msgbox("ov.ActiveDocument.ActiveWindow.Document.Frameset" & ChrW(10) & ex.Message)
                    '    End Try

                    'Else


                    '    Try
                    '        ov.ActiveDocument.ActiveWindow.View.SplitSpecial = Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneRevisionsVert
                    '    Catch ex As Exception
                    '        var1 = ex.Message
                    '        var1 = var1
                    '        'msgbox("ov.ActiveDocument.ActiveWindow.View.SplitSpecial" & ChrW(10) & ex.Message)
                    '    End Try

                    '    Try
                    '        ov.ActiveDocument.ActiveWindow.ShowSourceDocuments = Microsoft.Office.Interop.Word.WdShowSourceDocuments.wdShowSourceDocumentsBoth
                    '    Catch ex As Exception
                    '        'MsgBox("ov.ActiveDocument.ActiveWindow.ShowSourceDocuments" & ChrW(10) & ex.Message)
                    '        var1 = ex.Message
                    '        var1 = var1
                    '    End Try

                    '    If Count1 > 0 Then
                    '        Try
                    '            With ov.ActiveDocument.ActiveWindow.Document.Frameset.ChildFramesetItem(2)
                    '                .WidthType = Microsoft.Office.Interop.Word.WdFramesetSizeType.wdFramesetSizeTypePercent
                    '                .Width = 85
                    '            End With
                    '        Catch ex As Exception
                    '            var1 = ex.Message
                    '            var1 = var1
                    '            'MsgBox("ov.ActiveDocument.ActiveWindow.Document.Frameset 1" & ChrW(10) & ex.Message)
                    '        End Try

                    '        Try
                    '            With ov.ActiveDocument.ActiveWindow.Document.Frameset.ChildFramesetItem(2).ChildFramesetItem(1)
                    '                .WidthType = Microsoft.Office.Interop.Word.WdFramesetSizeType.wdFramesetSizeTypePercent
                    '                .Width = 42.5
                    '            End With
                    '        Catch ex As Exception
                    '            var1 = ex.Message
                    '            var1 = var1
                    '            ' MsgBox("ov.ActiveDocument.ActiveWindow.Document.Frameset 2" & ChrW(10) & ex.Message)
                    '        End Try
                    '    End If

                    'End If

                Catch ex As Exception

                End Try
            End If

        Next

        'make print view
        Try
            ov.ActiveDocument.ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
        Catch ex As Exception

        End Try


        If boolInVis Then
        Else
            If boolTemplate Then
            Else
                ov.Visible = True
            End If

        End If

        'Call DoReviewPane()


        'ov.Visible = True

    End Sub

    Sub CompareCode()

        'View side-by-side
        'Windows.CompareSideBySideWith "Document1"

        'Compare
        'Application.CompareDocuments(OriginalDocument:=Documents( _
        '"LIQ-000032-001-01.docx"), RevisedDocument:=Documents( _
        '"LIQ-000032-001-02.docx"), Destination:=wdCompareDestinationNew, _
        'Granularity:=wdGranularityWordLevel, CompareFormatting:=True, _
        'CompareCaseChanges:=True, CompareWhitespace:=True, CompareTables:=True, _
        'CompareHeaders:=True, CompareFootnotes:=True, CompareTextboxes:=True, _
        'CompareFields:=True, CompareComments:=True, CompareMoves:=True, _
        'RevisedAuthor:="LEE", IgnoreAllComparisonWarnings:=False)
        'ActiveWindow.ShowSourceDocuments = wdShowSourceDocumentsBoth


        'Dim objDoc1 As Word.Document
        'Dim objDoc2 As Word.Document

        'objDoc1 = Documents.Add
        'objDoc2 = Documents.Add

        'objDoc2.Activate()
        'objDoc2.Windows.CompareSideBySideWith objDoc1
        'Windows.ResetPositionsSideBySide()



        'https://msdn.microsoft.com/en-us/library/office/hh128820(v=office.14).aspx

        '' Save the current document, including this code.
        'Const path1 As String = "C:\Temp\Doc1.docm"
        'Const path2 As String = "C:\Temp\Doc2.docm"
        'Const path3 As String = "C:\Temp\Doc3.docm"

        'Dim doc1 As Document
        'Dim doc2 As Document
        'Dim doc3 As Document

        '' Save with macros enabled, because this code exists within
        '' the document that you are saving. If the document did not
        '' contain code, you would not need to specify the file format.
        'doc1 = ActiveDocument
        'doc1.SaveAs(path1, wdFormatXMLDocumentMacroEnabled)

        '' Make some changes to the current document.
        'doc2 = ActiveDocument
        'doc2.ApplyQuickStyleSet2 "Elegant"
        'ChangeTheDocument doc2

        '' Save the document as Doc2
        'doc2.SaveAs2(path2, wdFormatFlatXMLMacroEnabled)

        '' Open the original document
        'doc1 = Documents.Open(path1)

        'doc3 = Application.CompareDocuments(doc1, doc2, _
        ' Destination:=wdCompareDestinationNew, _
        ' Granularity:=wdGranularityWordLevel, _
        ' CompareFormatting:=True, _
        ' CompareCaseChanges:=True, _
        ' CompareWhiteSpace:=True)

        'doc3.SaveAs2(path3, wdFormatFlatXMLMacroEnabled)



    End Sub

    Private Sub chkRPane_CheckedChanged(sender As Object, e As EventArgs) Handles chkRPane.CheckedChanged

        'Call DoReviewPane()

    End Sub

    Private Sub chkRPane1_CheckedChanged(sender As Object, e As EventArgs) Handles chkRPane1.CheckedChanged

        'Call DoReviewPane()

    End Sub


    Private Sub cmdInsertDocument_Click(sender As Object, e As EventArgs) Handles cmdInsertDocument.Click

        'This button will allow users to insert a Word document into the Final Report

        Dim strM As String
        strM = "Under construction..."
        MsgBox(strM, vbInformation, "Under construction...")

    End Sub

    Function ShowOptionsForm(strSource As String) As Boolean

        ShowOptionsForm = False

        If boolReportGenAdvPrompt Then

            Dim frm As New frmFinalReportPrintOptions

            If StrComp(strSource, "Word", CompareMethod.Text) = 0 Then
                frm.gbChoice.Visible = True
            Else
                frm.gbChoice.Visible = False
            End If

            frm.ShowDialog()

            If frm.boolCancel Then
                ShowOptionsForm = False
            Else
                ShowOptionsForm = True
            End If

            boolAsIsOp = frm.rbAsIs.Checked
            boolReadOnlyOp = frm.rbReadOnly.Checked
            boolNoneOp = frm.rbNone.Checked
            boolWaterMarkOp = frm.rbWaterMark.Checked
            boolTextOp = frm.rbText.Checked
            boolCenterOp = frm.rbCenter.Checked
            boolFooterOp = frm.rbFooter.Checked
            boolDTCreatedOp = frm.chkDTCreated.Checked
            boolDTReportedOp = frm.chkDTReported.Checked
            boolDocIDOp = frm.chkID.Checked
            boolDocGenOp = frm.chkGenerator.Checked
            boolDocOwnerOp = frm.chkOwner.Checked

            frm.Dispose()

        Else

            boolAsIsOp = True
            boolReadOnlyOp = False
            boolNoneOp = True
            ShowOptionsForm = True

        End If

    End Function

    Function ESigPrompt() As Boolean

        ESigPrompt = False

        If gboolAuditTrail And gboolESig Then

            Dim frm As New frmESig

            frm.ShowDialog()

            If frm.boolCancel Then
                frm.Dispose()
                GoTo end2
            End If

            frm.Dispose()

            ESigPrompt = True

        Else
            ESigPrompt = True
        End If

end2:

    End Function

    Private Sub cmdPrint_Click(sender As Object, e As EventArgs) Handles cmdPrint.Click

        Dim strM As String

        strM = "User does not have permission to print this document."
        If BOOLALLOWFINALREPORTPRINT Then
        Else
            MsgBox(strM, vbInformation, "Invalid action...")
            GoTo end1
        End If

        If Me.panSC1.Visible Then
            If Me.ovDC1.IsOpened Then
            Else
                strM = "A document must be loaded."
                MsgBox(strM, vbInformation, "Invalid action...")
                GoTo end1
            End If
        Else
            If Me.ovDC2.IsOpened Then
            Else
                strM = "A document must be loaded."
                MsgBox(strM, vbInformation, "Invalid action...")
                GoTo end1
            End If
        End If

        If ShowOptionsForm("Print") Then
        Else
            Exit Sub
        End If

        If ESigPrompt() Then
        Else
            GoTo end2
        End If

        boolNoneOp = True

        If Me.panSC1.Visible Then
            If Me.ovDC1.IsOpened Then
                Me.ovDC1.PrintDialog()
            End If
        Else
            If Me.ovDC2.IsOpened Then
                Me.ovDC.PrintDialog()
            End If
        End If

end1:

end2:

    End Sub

    Sub DoEdit(strE As String)

        'Me.gbSaveType.Enabled = True

        Dim strM As String
        Dim var1

        Dim doc As Microsoft.Office.Interop.Word.Document
        doc = Me.ovDC1.ActiveDocument
        Dim vP

        Dim strP As String
        Dim strP1 As String

        If boolHasBeenSaved Then
            strP = GetPassword(gid1)
            strP1 = PasswordUnEncrypt(strP)
        Else
            strP1 = PasswordUnEncrypt(tPswd)
            If Len(strP1) = 0 Then
                Try
                    strP = GetPassword(gid1)
                    strP1 = PasswordUnEncrypt(strP)
                    var1 = var1
                Catch ex As Exception
                    var1 = var1
                End Try
            End If
        End If


        Dim intR As Int16
        Try
            intR = UBound(arrRanges1)
        Catch ex As Exception
            intR = 1000
        End Try

        Select Case strE
            Case "Edit"

                vP = doc.ProtectionType
                '20181110 LEE:
                'if vp = 3 (readonly) then must unprotect doc before unprotecting ov (DoWrite)
                If vP = 3 And Len(strP1) <> 0 Then
                    Try
                        doc.Unprotect(strP1)
                    Catch ex As Exception
                        var1 = var1
                    End Try

                End If

                'first unprotect ov
                Call DoWrite()
                'if actual word doc has protected ranges, protection type will stay <> 5

            Case "Save"


            Case "Cancel"


        End Select



        vP = doc.ProtectionType ' 1 - 5 or -1 for an error. NoProtection = 5

        Try
            'If doc.ProtectionType = WdProtectionType.wdNoProtection Then
            If intR = 1000 Then

                Select Case strE
                    Case "Edit"

                        vP = doc.ProtectionType
                        '20181110 LEE:
                        'if vp = 3 (readonly) then must unprotect doc before unprotecting ov (DoWrite)
                        If vP = 3 And Len(strP1) <> 0 Then
                            Try
                                doc.Unprotect(strP1)
                            Catch ex As Exception
                                var1 = var1
                            End Try

                        End If

                        Call DoWrite()

                    Case "Save"

                        Call DoReadOnly(Me.ovDC1)

                    Case "Cancel"

                        Call DoReadOnly(Me.ovDC1)

                End Select

            Else

                Select Case strE
                    Case "Edit"

                        Call RestoreProtectedRanges(Me.ovDC1, gid1)
                        'MsgBox("DoEdit:Edit:RestoreProtectedRanges")
                        '20180919 LEE: Why is DoReadOnly being called? Doc should be editable
                        'Call DoReadOnly(Me.ovDC1)
                        'MsgBox("DoEdit:Edit:DoReadOnly")
                    Case "Save"

                        Call StoreProtectedRanges(Me.ovDC1, gid1)

                        '20181110 LEE:
                        'Shouldn't need to do any of this

                        'doc = Me.ovDC1.ActiveDocument
                        'vP = doc.ProtectionType
                        'If vP = 3 Then
                        '    Try
                        '        Call DoReadOnly(Me.ovDC1)
                        '    Catch ex As Exception
                        '        var1 = var1
                        '    End Try
                        'Else
                        '    '20181110 LEE:
                        '    'protect doc
                        '    Try
                        '        doc.Protect(Password:=strP1, NoReset:=False, Type:=Microsoft.Office.Interop.Word.WdProtectionType.wdAllowOnlyReading, UseIRM:=False, EnforceStyleLock:=False)
                        '    Catch ex As Exception
                        '        var1 = var1
                        '    End Try
                        '    Call DoReadOnly(Me.ovDC1)
                        '    var1 = var1
                        'End If

                    Case "Cancel"

                End Select

            End If
        Catch ex As Exception

        End Try

end1:

    End Sub

    Function DefaultFormats(strE As String) As Boolean

        DefaultFormats = False

        'Me.gbSaveType.Enabled = True

        Dim strM As String

        Select Case strE
            Case "Edit"

                If BOOLEDITFINALREPORT Then
                Else
                    strM = "User does not have permissions to Edit reports."
                    'strM = strM & ChrW(10) & "Therefore, user may not execute a Save action."
                    MsgBox(strM, vbInformation, "Invalid action...")
                    GoTo end1
                End If

                Me.cmdEdit.Enabled = False
                Me.cmdEdit.BackColor = System.Drawing.Color.Gray

                Me.cmdSave.Enabled = True
                Me.cmdSave.BackColor = System.Drawing.Color.Gainsboro

                Me.cmdCancel.Enabled = True
                Me.cmdCancel.BackColor = System.Drawing.Color.Gainsboro

                Me.panOpen.Enabled = False
                'Me.cmdEdit.BackColor = System.Drawing.Color.Gray

                'Me.gbSaveType.Enabled = False
                'Me.cmdEdit.BackColor = System.Drawing.Color.Gray

                Me.gbLoad.Enabled = False
                'Me.cmdEdit.BackColor = System.Drawing.Color.Gray

                Me.cmdClearCompare.Enabled = False
                Me.cmdClearCompare.BackColor = System.Drawing.Color.Gray

                Me.cmdCompareFinalReport.Enabled = False
                Me.cmdCompareFinalReport.BackColor = System.Drawing.Color.Gray

                Me.cmdCompareSection.Enabled = False
                Me.cmdCompareSection.BackColor = System.Drawing.Color.Gray

                Me.cmdExit.Enabled = False
                Me.cmdExit.BackColor = System.Drawing.Color.Gray

                Me.cmdWord.BackColor = System.Drawing.Color.Gray
                Me.cmdOpenPDF.BackColor = System.Drawing.Color.Gray
                Me.cmdPrint.BackColor = System.Drawing.Color.Gray
                Me.cmdInsertDocument.BackColor = System.Drawing.Color.Gray

                Me.dgvFinalReports.Enabled = False
                Me.dgvFinalReports.DefaultCellStyle.BackColor = Color.Gray
                Me.dgvSections.Enabled = False
                Me.dgvSections.DefaultCellStyle.BackColor = Color.Gray

                Me.cmdPaste.Enabled = True
                Me.cmdPaste.BackColor = System.Drawing.Color.Gainsboro

                Me.cmdCopy.Enabled = True
                Me.cmdCopy.BackColor = System.Drawing.Color.Gainsboro

                'set focus to a Save
                Me.cmdSave.Focus()

            Case "Save", "Cancel"

                'Call StoreProtectedRanges(Me.ovDC1, gid1)
                If BOOLFINALREPORTLOCKED Then
                Else
                    Me.cmdEdit.Enabled = True
                    Me.cmdEdit.BackColor = System.Drawing.Color.Gainsboro
                End If

                Me.cmdSave.Enabled = False
                Me.cmdSave.BackColor = System.Drawing.Color.Gray

                Me.cmdCancel.Enabled = False
                Me.cmdCancel.BackColor = System.Drawing.Color.Gray

                Me.panOpen.Enabled = True
                'Me.cmdSave.BackColor = System.Drawing.Color.Gainsboro

                'Me.gbSaveType.Enabled = True
                'Me.cmdSave.BackColor = System.Drawing.Color.Gainsboro

                Me.gbLoad.Enabled = True
                'Me.cmdSave.BackColor = System.Drawing.Color.Gainsboro

                Me.cmdClearCompare.Enabled = True
                Me.cmdClearCompare.BackColor = System.Drawing.Color.Gainsboro

                Me.cmdCompareFinalReport.Enabled = True
                Me.cmdCompareFinalReport.BackColor = System.Drawing.Color.Gainsboro

                Me.cmdCompareSection.Enabled = True
                Me.cmdCompareSection.BackColor = System.Drawing.Color.Gainsboro

                Me.cmdExit.Enabled = True
                Me.cmdExit.BackColor = System.Drawing.Color.Gainsboro

                Me.cmdWord.BackColor = System.Drawing.Color.Gainsboro
                Me.cmdOpenPDF.BackColor = System.Drawing.Color.Gainsboro
                Me.cmdPrint.BackColor = System.Drawing.Color.Gainsboro
                Me.cmdInsertDocument.BackColor = System.Drawing.Color.Gainsboro

                Me.dgvFinalReports.Enabled = True
                Me.dgvFinalReports.DefaultCellStyle.BackColor = Color.Empty 'resets to default

                Me.dgvSections.Enabled = True
                Me.dgvSections.DefaultCellStyle.BackColor = Color.Empty 'resets to default

                Me.cmdPaste.Enabled = False
                Me.cmdPaste.BackColor = System.Drawing.Color.Gray

                Me.cmdCopy.Enabled = False
                Me.cmdCopy.BackColor = System.Drawing.Color.Gray

                'set focus to a Edit
                Me.cmdEdit.Focus()

        End Select

        Call SpecialFormats()

        DefaultFormats = True

end1:

    End Function

    Sub WordCompareFormats()

        Me.cmdEdit.Enabled = False
        Me.cmdEdit.BackColor = System.Drawing.Color.Gray

        Me.cmdSave.Enabled = False
        Me.cmdSave.BackColor = System.Drawing.Color.Gray

        Me.cmdCancel.Enabled = False
        Me.cmdCancel.BackColor = System.Drawing.Color.Gray

        Me.panOpen.Enabled = False
        'Me.cmdEdit.BackColor = System.Drawing.Color.Gray

        'Me.gbSaveType.Enabled = False
        'Me.cmdEdit.BackColor = System.Drawing.Color.Gray

        'Me.gbLoad.Enabled = False
        ''Me.cmdEdit.BackColor = System.Drawing.Color.Gray

        'Me.cmdCompareFinalReport.Enabled = False
        'Me.cmdCompareFinalReport.BackColor = System.Drawing.Color.Gray

        'Me.cmdCompareSection.Enabled = False
        'Me.cmdCompareSection.BackColor = System.Drawing.Color.Gray

        'Me.cmdExit.Enabled = False
        'Me.cmdExit.BackColor = System.Drawing.Color.Gray

        Me.cmdWord.BackColor = System.Drawing.Color.Gray
        Me.cmdOpenPDF.BackColor = System.Drawing.Color.Gray
        Me.cmdPrint.BackColor = System.Drawing.Color.Gray
        Me.cmdInsertDocument.BackColor = System.Drawing.Color.Gray

        'Me.dgvFinalReports.Enabled = False
        'Me.dgvFinalReports.DefaultCellStyle.BackColor = Color.Gray
        'Me.dgvSections.Enabled = False
        'Me.dgvSections.DefaultCellStyle.BackColor = Color.Gray


    End Sub

    Sub InitGeneratedReport()

        If gboolER Then
        Else
            GoTo end1
        End If

        If StrComp(strPrevForm, "ReportHistory", CompareMethod.Text) = 0 Then
            GoTo end1
        End If

        Call StoreProtectedRanges(Me.ovDC1, gid1)

        Me.cmdEdit.Enabled = False
        Me.cmdEdit.BackColor = System.Drawing.Color.Gray

        Me.cmdSave.Enabled = True
        Me.cmdSave.BackColor = System.Drawing.Color.Gainsboro

        Me.cmdCancel.Enabled = False
        Me.cmdCancel.BackColor = System.Drawing.Color.Gainsboro

        Me.cmdCancel.Enabled = False
        Me.cmdCancel.BackColor = System.Drawing.Color.Gray

        Me.panOpen.Enabled = False
        'Me.cmdEdit.BackColor = System.Drawing.Color.Gray

        'Me.gbSaveType.Enabled = False
        'Me.cmdEdit.BackColor = System.Drawing.Color.Gray

        Me.gbLoad.Enabled = False
        Me.gbLoad.Visible = True
        'Me.cmdEdit.BackColor = System.Drawing.Color.Gray

        Me.cmdClearCompare.Enabled = False
        Me.cmdClearCompare.BackColor = System.Drawing.Color.Gray

        Me.cmdCompareFinalReport.Enabled = False
        Me.cmdCompareFinalReport.BackColor = System.Drawing.Color.Gray

        Me.cmdCompareSection.Enabled = False
        Me.cmdCompareSection.BackColor = System.Drawing.Color.Gray

        'Me.cmdExit.Enabled = False
        'Me.cmdExit.BackColor = System.Drawing.Color.Gray

        Me.cmdWord.BackColor = System.Drawing.Color.Gray
        Me.cmdOpenPDF.BackColor = System.Drawing.Color.Gray
        Me.cmdPrint.BackColor = System.Drawing.Color.Gray
        Me.cmdInsertDocument.BackColor = System.Drawing.Color.Gray

        Me.dgvFinalReports.Enabled = False
        Me.dgvFinalReports.DefaultCellStyle.BackColor = Color.Gray
        Me.dgvSections.Enabled = False
        Me.dgvSections.DefaultCellStyle.BackColor = Color.Gray

        Me.cmdPaste.Enabled = False
        Me.cmdPaste.BackColor = System.Drawing.Color.Gray

        Me.cmdCopy.Enabled = False
        Me.cmdCopy.BackColor = System.Drawing.Color.Gray

        Call SpecialFormats()

end1:

    End Sub

    Sub SpecialFormats()

        'special cmdEdit, cmdSave, cmdPrint
        If BOOLFORCEFINALREPORTPDF Then

            Me.cmdPrint.Enabled = False
            Me.cmdPrint.BackColor = System.Drawing.Color.Gray

            Me.cmdWord.Enabled = False
            Me.cmdWord.BackColor = System.Drawing.Color.Gray

        Else

        End If

        If Me.cmdSave.Enabled Then
        Else
            If Me.ovDC1.IsOpened Then
                If Me.rbCompare.Checked And Me.rbCompare.Checked Then
                    'cmdEdit status governed by WordCompareFormats
                Else
                    If BOOLFINALREPORTLOCKED Then
                    Else
                        Me.cmdEdit.Enabled = True
                        Me.cmdEdit.BackColor = System.Drawing.Color.Gainsboro
                    End If
                  
                End If

            Else
                Me.cmdEdit.Enabled = False
                Me.cmdEdit.BackColor = System.Drawing.Color.Gray

                Me.cmdClearCompare.Enabled = False
                Me.cmdClearCompare.BackColor = System.Drawing.Color.Gray
            End If
        End If


    End Sub

    Function GetPassword(id As Int64) As String

        GetPassword = ""

        Dim strP As String

        Dim strF As String
        Dim rows() As DataRow
        Dim dtbl As System.Data.DataTable

        'first check dv1
        Dim dv1 As DataView = Me.dgvFinalReports.DataSource
        dtbl = dv1.ToTable
        strF = "ID_TBLFINALREPORT = " & id
        rows = dtbl.Select(strF)
        If rows.Length = 0 Then
            'check dv2
            Dim dv2 As DataView = Me.dgvSections.DataSource
            dtbl = dv2.ToTable
            rows = dtbl.Select(strF)
            If rows.Length = 0 Then
                GoTo end1
            Else
                strP = NZ(rows(0).Item("CHARPASSWORD"), "")
            End If
        Else
            strP = NZ(rows(0).Item("CHARPASSWORD"), "")
        End If

        GetPassword = strP

end1:

    End Function


    Sub StoreProtectedRanges(ov As AxEDOfficeLib.AxEDOffice, id As Int64)

        Dim boolRecord As Boolean
        Dim var1, var2, var3
        Dim vP

        If StrComp(ov.Name, "OVDC1", CompareMethod.Text) = 0 Then
            boolRecord = True
        Else
            boolRecord = False
        End If

        Cursor.Current = Cursors.WaitCursor

        Dim strP As String ' = GetPassword(id)
        Dim strP1 As String ' = PasswordUnEncrypt(strP)

        If boolHasBeenSaved Then
            strP = GetPassword(id)
            strP1 = PasswordUnEncrypt(strP)
        Else
            strP1 = PasswordUnEncrypt(tPswd)
        End If

        'must do this:
        '1. Go through and record editable ranges
        '2. Unprotect with stored password
        '3. Do in code equivalent of unchecking Word-Review-AllowOnlyThisTypeOfEditing
        '4. Do OV Readonly

        'unlock document
        Dim doc As Microsoft.Office.Interop.Word.Document
        Try
            doc = ov.ActiveDocument
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
            GoTo end1
        End Try


        Try
            If doc.ProtectionType = Microsoft.Office.Interop.Word.WdProtectionType.wdNoProtection Then
                ReDim arrRanges1(1000)
                Call DoReadOnly(ov)
                GoTo end1
            End If
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
            GoTo end1
        End Try

        'ReDim arrRanges1(0)

        Try

            '1. Go through and record editable ranges
            If boolRecord Then
                ReDim arrRanges1(1000)
            End If

            Dim int1 As Integer
            Dim int2 As Integer
            Dim intPT As Integer
            Dim Count1 As Integer
            Dim intPos1 As Long
            Dim intPos2 As Long

            intPos1 = 99

            Dim boolA As Boolean

            intPT = doc.ProtectionType

            Dim rng1 As Microsoft.Office.Interop.Word.Range ' Word.Range
            Dim rng2 As Microsoft.Office.Interop.Word.Range 'Word.Range
            Dim intR As Int16

            doc.Application.ScreenUpdating = False

            doc.Application.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

            Call RemoveProtectedRanges(ov, boolRecord, id)

            vP = doc.ProtectionType
            If vP = 3 Then 'read only
            Else

                If Len(strP1) <> 0 Then
                    Try
                        '2. Unprotect with stored password
                        doc.Unprotect(strP1)
                    Catch ex As Exception
                        var1 = ex.Message
                        var1 = var1
                    End Try

                    Try
                        doc.Protect(Microsoft.Office.Interop.Word.WdProtectionType.wdNoProtection)
                    Catch ex As Exception
                        var1 = ex.Message
                        var1 = var1
                    End Try
                End If
        

                '4. Do OV Readonly
                Call DoReadOnly(ov)

            End If

        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

        doc.Application.ScreenUpdating = True

end1:

        Cursor.Current = Cursors.Default

    End Sub

    Sub JustRemoveProtectedRanges(ov As AxEDOfficeLib.AxEDOffice, id As Int64)

        Try

            Dim strP As String = GetPassword(id)
            Dim strP1 As String = PasswordUnEncrypt(strP)

            Dim int1 As Integer
            Dim int2 As Integer
            Dim intPT As Integer
            Dim Count1 As Integer
            Dim intPos1 As Long
            Dim intPos2 As Long
            Dim var1

            Dim rng1 As Microsoft.Office.Interop.Word.Range ' Word.Range
            Dim rng2 As Microsoft.Office.Interop.Word.Range 'Word.Range
            Dim intR As Int16

            Dim doc As Microsoft.Office.Interop.Word.Document
            doc = ov.ActiveDocument

            'ReDim arrR(0)

            If doc.ProtectionType = Microsoft.Office.Interop.Word.WdProtectionType.wdNoProtection Then
                GoTo end1
            End If

            doc.Application.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

            For Count1 = 1 To 1000

                'Note: will go to 2nd region first
                'the last found region will be the first region at the beginning of the document
                'if Selection.HomeKey Unit:=wdStory is called first

                doc.Application.Selection.GoToEditableRange(Microsoft.Office.Interop.Word.WdEditorType.wdEditorEveryone)
                rng1 = doc.Application.Selection.Range

                If intR = 1 Then
                    intPos1 = rng1.Start
                Else
                    intPos2 = rng1.Start
                    If intPos1 = intPos2 Then
                        Exit For
                    End If
                End If

                intR = intR + 1

                If rng1.Editors.Item(1).Name = "Everyone" Then
                    rng1.Editors.Item(1).Delete()
                End If

            Next

            doc.Application.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

            If Len(strP1) = 0 Then
            Else
                Try
                    '2. Unprotect with stored password
                    doc.Unprotect(strP1)
                Catch ex As Exception
                    var1 = ex.Message
                    var1 = var1
                End Try
            End If
      

            Try
                doc.Protect(Microsoft.Office.Interop.Word.WdProtectionType.wdNoProtection)
            Catch ex As Exception
                var1 = ex.Message
                var1 = var1
            End Try

            'save it 
            doc.Save()


        Catch ex As Exception

        End Try

end1:

    End Sub

    Sub RemoveProtectedRanges(ov As AxEDOfficeLib.AxEDOffice, boolRecord As Boolean, id As Int64)

        Dim var1, var2, var3
        Dim arrR() As Microsoft.Office.Interop.Word.Range

        Dim strM As String

        Dim strP As String ' = GetPassword(id)
        Dim strP1 As String ' = PasswordUnEncrypt(strP)

        If boolHasBeenSaved Then
            strP = GetPassword(id)
            strP1 = PasswordUnEncrypt(strP)
        Else
            strP1 = PasswordUnEncrypt(tPswd)
        End If

        'must do this:
        '1. Go through and record editable ranges
        '2. Unprotect with stored password
        '3. Do in code equivalent of unchecking Word-Review-AllowOnlyThisTypeOfEditing
        '4. Do OV Readonly

        'unlock document
        Dim doc As Microsoft.Office.Interop.Word.Document
        doc = ov.ActiveDocument

        'ReDim arrR(0)

        If doc.ProtectionType = Microsoft.Office.Interop.Word.WdProtectionType.wdNoProtection Then
            GoTo end1
        End If

        Try

            '1. Go through and record editable ranges
            ReDim arrR(1000)

            Dim int1 As Integer
            Dim int2 As Integer
            Dim intPT As Integer
            Dim Count1 As Integer
            Dim intPos1 As Long
            Dim intPos2 As Long

            intPos1 = 99

            Dim boolA As Boolean

            intPT = doc.ProtectionType

            Dim rng1 As Microsoft.Office.Interop.Word.Range ' Word.Range
            Dim rng2 As Microsoft.Office.Interop.Word.Range 'Word.Range
            Dim intR As Int16

            doc.Application.ScreenUpdating = False

            doc.Application.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

            Try
                intR = 0
                For Count1 = 1 To 100

                    'Note: action will go to 2nd region first
                    'the last found region will be the first region at the beginning of the document
                    'if Selection.HomeKey Unit:=wdStory is called first

                    doc.Application.Selection.GoToEditableRange(Microsoft.Office.Interop.Word.WdEditorType.wdEditorEveryone)
                    rng1 = doc.Application.Selection.Range

                    If intR = 1 Then
                        intPos1 = rng1.Start
                    Else
                        intPos2 = rng1.Start
                        If intPos1 = intPos2 Then
                            Exit For
                        End If
                    End If

                    ''debug
                    'strM = "intPos1: " & intPos1 & ChrW(10) & "intPos2: " & intPos2 & ChrW(10) & "Count1: " & Count1
                    'MsgBox(strM)

                    Try
                        If rng1.Editors.Item(1).Name = "Everyone" Then
                            rng1.Editors.Item(1).Delete()
                        End If
                    Catch ex As Exception
                        If intPos1 = 0 Or intPos2 = 0 Then
                            Exit For
                        End If
                        GoTo next1
                    End Try

                    intR = intR + 1

                    If intR > UBound(arrR) Then
                        ReDim Preserve arrR(UBound(arrR) + 100)
                    End If
                    arrR(intR) = rng1

                    If intPos1 = 0 Or intPos2 = 0 Then
                        Exit For
                    End If

next1:

                Next

                If boolRecord Then
                    ReDim Preserve arrR(intR)
                    ReDim arrRanges1(intR)
                    arrRanges1 = arrR
                End If

                'ReDim Preserve arrRanges1(intR)

            Catch ex As Exception
                var1 = ex.Message
                var1 = var1
            End Try

            doc.Application.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

end1:

    End Sub

    Sub RestoreProtectedRanges(ov As AxEDOfficeLib.AxEDOffice, id As Int64)


        If StrComp(ov.Name, "OVDC1", CompareMethod.Text) = 0 Then
        Else
            GoTo end1
        End If

        Dim intR As Int16 = UBound(arrRanges1)

        If intR = 1000 Or intR = 0 Then
            'Call DoWrite()
            GoTo end1
        End If

        Cursor.Current = Cursors.WaitCursor

        '20181110 LEE:
        'The next line is a problem if document has protected ranges
        'Call DoWrite()

        'now re-establish ranges

        Dim strP As String ' = GetPassword(id)
        Dim strP1 As String ' = PasswordUnEncrypt(strP)

        If boolHasBeenSaved Then
            strP = GetPassword(id)
            strP1 = PasswordUnEncrypt(strP)
        Else
            strP1 = PasswordUnEncrypt(tPswd)
        End If

        Dim var1, var2, var3

        Dim rng1 As Microsoft.Office.Interop.Word.Range ' Word.Range
        Dim rng2 As Microsoft.Office.Interop.Word.Range 'Word.Range


        Dim doc As Microsoft.Office.Interop.Word.Document
        doc = ov.ActiveDocument
        doc.Application.ScreenUpdating = False

        Dim Count1 As Int16

        var1 = doc.ProtectionType

        If doc.ProtectionType = Microsoft.Office.Interop.Word.WdProtectionType.wdNoProtection Then
            'doc.Protect(WdProtectionType.wdAllowOnlyReading)
        Else
            Try
                'doc.Unprotect(strP1)
            Catch ex As Exception
                var1 = var1
            End Try

            'doc.Protect(WdProtectionType.wdAllowOnlyReading)
        End If
        'doc.Protect(WdProtectionType.wdAllowOnlyReading)

        Try
            For Count1 = 1 To intR

                'debug
                var2 = UBound(arrRanges1)

                var1 = arrRanges1(Count1)
                rng1 = var1 ' arrRanges1(Count1)
                rng1.Select()

                'protect this
                'Selection.Editors.Add wdEditorEveryone
                rng1.Editors.Add(Microsoft.Office.Interop.Word.WdEditorType.wdEditorEveryone)

            Next
        Catch ex As Exception
            var1 = ex.Message
            'MsgBox(var1 & ChrW(10) & "Ubound: " & UBound(arrRanges1))
        End Try


        Try
            If doc.ProtectionType = Microsoft.Office.Interop.Word.WdProtectionType.wdAllowOnlyReading Then
            Else
                doc.Protect(Password:=strP1, NoReset:=False, Type:=Microsoft.Office.Interop.Word.WdProtectionType.wdAllowOnlyReading, UseIRM:=False, EnforceStyleLock:=False)
            End If

        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

        Call HilightSections()

        doc.Application.ScreenUpdating = True

end1:

        Cursor.Current = Cursors.Default


    End Sub

    Private Sub chkShowHilite_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowHilite.CheckedChanged

        Call HilightSections()

    End Sub

    Sub HilightSections()

        Dim doc As Microsoft.Office.Interop.Word.Document

        Dim var1
        Dim boolH As Boolean = False

        Try
            doc = Me.ovDC1.ActiveDocument

            If Me.chkShowHilite.Checked Then
                doc.Windows(1).View.ShadeEditableRanges = True
            Else
                doc.Windows(1).View.ShadeEditableRanges = False
            End If

            If doc.ProtectionType = Microsoft.Office.Interop.Word.WdProtectionType.wdNoProtection Then
            Else
                boolH = True
            End If

        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

        Try
            doc = Me.ovDC2.ActiveDocument

            If Me.chkShowHilite.Checked Then
                doc.Windows(1).View.ShadeEditableRanges = True
            Else
                doc.Windows(1).View.ShadeEditableRanges = False
            End If

            If doc.ProtectionType = Microsoft.Office.Interop.Word.WdProtectionType.wdNoProtection Then
            Else
                boolH = True
            End If

        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

        Try
            Call ShowProtDocStuff(boolH)
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

    End Sub

    Sub ShowProtDocStuff(boolH As Boolean)

        Me.chkShowHilite.Visible = boolH

    End Sub




    Private Sub cmdClearCompare_Click(sender As Object, e As EventArgs) Handles cmdClearCompare.Click

        Try
            If Me.ovDC2.IsOpened Then
            Else
                GoTo end1
            End If
        Catch ex As Exception
            GoTo end1
        End Try

        boolFromClearCompare = True

        Me.lblProgress1.Text = "Clearing compare..."
        Me.panProgress1.Visible = True
        Me.panProgress1.Refresh()

        Me.txtComparedDocDescription.Text = ""

        '20190115 LEE
        strPathT2 = ""

        Try
            Me.ovDC2.CloseDoc(False)
        Catch ex As Exception

        End Try

        Me.rbLoad.Checked = True

        'must choose correct entry in table


        Try
            Call SizePanes(False)
        Catch ex As Exception

        End Try

        Me.panProgress1.Visible = False
        Me.panProgress1.Refresh()
        Me.lblProgress1.Text = ""

        boolFromClearCompare = False

        Dim str1 As String
        str1 = "gid1: " & gid1 & ChrW(10) & "gid2: " & gid2
        'MsgBox(str1)

        '20180919 LEE:
        'must select gid1 in either final or section
        Dim dgv1 As DataGridView = Me.dgvFinalReports
        Dim dgv2 As DataGridView = Me.dgvSections

        Dim dv1 As DataView = dgv1.DataSource
        Dim dv2 As DataView = dgv2.DataSource

        dv1.Sort = "ID_TBLFINALREPORT DESC"
        dv2.Sort = "ID_TBLFINALREPORT DESC"

        Dim i1 As Int16
        Dim i2 As Int16

        i1 = dv1.Find(gid1)
        i2 = dv2.Find(gid1)

        'MsgBox("i1: " & i1 & ChrW(10) & "i2: " & i2)

        'currentcell
        If i1 >= 0 Then
            'make selection in Final Reports
            dgv1.CurrentCell = dgv1.Item(GetVisibleCol(dgv1), i1)
            dgv1.CurrentRow.Selected = True
        Else
            If i2 >= 0 Then
                dgv2.CurrentCell = dgv2.Item(GetVisibleCol(dgv2), i2)
                dgv2.CurrentRow.Selected = True
            End If
        End If

end1:

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Try
            Me.ovDC.ActiveDocument.ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
        Catch ex As Exception

        End Try
        Try
            Me.ovDC1.ActiveDocument.ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
        Catch ex As Exception

        End Try
        Try
            Me.ovDC2.ActiveDocument.ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
        Catch ex As Exception

        End Try


    End Sub

    Private Sub cmdCopy_Click(sender As Object, e As EventArgs) Handles cmdCopy.Click

        Call CopyFromCompare()

    End Sub

    Sub CopyFromCompare()

        Dim intR As Short
        Dim strM As String

        strM = "Ary you sure you want to paste the Compare Document selectiion to the Loaded Document?" & ChrW(10) & ChrW(10)
        strM = strM & "Ensure the cursor is placed in the Loaded Document to which the information is to be pasted."
        intR = MsgBox(strM, vbOKCancel, "Ensure...")
        If intR = 1 Then
        Else
            GoTo end1
        End If

        Dim boolComp As Boolean = False

        Call PasteFromCompare()

end1:

    End Sub


    Private Sub cmdPaste_Click(sender As Object, e As EventArgs) Handles cmdPaste.Click

        Dim intR As Short
        Dim strM As String

        strM = "Ary you sure you want to paste from the clipboard to the Loaded Document?" & ChrW(10) & ChrW(10)
        strM = strM & "Ensure the cursor is placed in the Loaded Document to which the information is to be pasted."
        intR = MsgBox(strM, vbOKCancel, "Ensure...")
        If intR = 1 Then
        Else
            GoTo end1
        End If

        Dim boolComp As Boolean = False

        Call PasteFromClipboard()

end1:

    End Sub

    Sub PasteFromClipboard()

        Dim intR As Short
        Dim strM As String

        Dim var1
        strM = ""
        Dim strM1 As String = ""

        Try

            Dim boolProt As Boolean = False
            Dim doc As Microsoft.Office.Interop.Word.Document
            Try
                If doc.ProtectionType = Microsoft.Office.Interop.Word.WdProtectionType.wdNoProtection Then
                Else
                    boolProt = True
                    strM1 = "This document has protected ranges. It is possible you are attempting to paste to a protected area."
                End If
            Catch ex As Exception
                strM1 = "This document has protected ranges. It is possible you are attempting to paste to a protected area."
            End Try

            Me.ovDC1.WordPasteFromClipboard()
        Catch ex As Exception
            var1 = ex.Message
            If Len(strM1) = 0 Then
                strM1 = "There was a problem pasting to this document."
            End If
            strM = strM1 & ChrW(10) & ChrW(10) & var1.ToString
            MsgBox(strM, vbInformation, "Problem pasting from clipboard...")
        End Try

end1:

    End Sub

    Sub PasteFromCompare()

        Dim strM As String = ""
        Dim strM1 As String = ""

        Dim var1

        Try
            If Me.ovDC2.IsOpened Then
            Else
                GoTo end1
            End If
        Catch ex As Exception
            strM1 = "There seems to be a problem obtaining information from the Compare Document."
            GoTo end1
        End Try



        Dim ov1 As AxEDOfficeLib.AxEDOffice = Me.ovDC1
        Dim ov2 As AxEDOfficeLib.AxEDOffice = Me.ovDC2

        Dim doc1 As Microsoft.Office.Interop.Word.Document = ov1.ActiveDocument
        Dim doc2 As Microsoft.Office.Interop.Word.Document = ov2.ActiveDocument

        Dim rng1 As Microsoft.Office.Interop.Word.Range
        Dim rng2 As Microsoft.Office.Interop.Word.Range

        Try
            rng2 = doc2.Application.Selection.Range
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
            If Len(strM1) = 0 Then
                strM1 = var1
            Else
                strM1 = strM1 & ChrW(10) & var1
            End If
        End Try

        Try
            rng1 = doc1.Application.Selection.Range
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
            If Len(strM1) = 0 Then
                strM1 = var1
            Else
                strM1 = strM1 & ChrW(10) & var1
            End If
        End Try

        Try
            Clipboard.Clear()
        Catch ex As Exception

        End Try

        Try
            rng2.Copy()
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
            If Len(strM1) = 0 Then
                strM1 = var1
            Else
                strM1 = strM1 & ChrW(10) & var1
            End If
        End Try

        Try
            rng1.Paste()
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
            If Len(strM1) = 0 Then
                strM1 = var1
            Else
                strM1 = strM1 & ChrW(10) & var1
            End If

            strM = "The destination to which the information is to be pasted seems to be protected in the Loaded Document."
            strM = strM & ChrW(10) & ChrW(10) & strM1

            MsgBox(strM, vbInformation, "Invalid action...")

        End Try

end1:

        Try
            Clipboard.Clear()
        Catch ex As Exception

        End Try

    End Sub

    Sub DocumentCompareToolTips()

        ' Create the ToolTip and associate with the Form container.
        Dim toolTip1 As New ToolTip()
        Dim str1 As String

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
            str1 = "Paste from Clipboard."
            toolTip1.SetToolTip(Me.cmdPaste, str1)

            str1 = "Copy Compare Document selection to cursor location in Loaded Document."
            toolTip1.SetToolTip(Me.cmdCopy, str1)

            str1 = "All documents are locked when Final Report is locked."
            toolTip1.SetToolTip(Me.lblFinalReportLocked, str1)

            'Grid
            'Me.dgvMaster.Columns.Item("CHARTYPE").ToolTipText = "Choose type of Appendix/Figure"
            'Me.dgvMaster.Columns.Item("CHARFCID").ToolTipText = "Choose a field code ID for Word Template"
            'Me.dgvMaster.Columns.Item("BOOLW").ToolTipText = "W: Insert figures from a word document"
            'str1 = "Enter Caption for the Appendix/Figure."
            'str1 = str1 & ChrW(10) & "Note that Caption is ignored if W* column is checked."
            'Me.dgvMaster.Columns.Item("CHARTITLE").ToolTipText = str1 '"Enter Title for the Appendix/Figure"
            'Me.dgvMaster.Columns.Item("CHARPATH").ToolTipText = "Choose directory path for supporting file"""
            'Me.dgvMaster.Columns.Item("INTORDER").ToolTipText = "A: Order in which appendices/figures are put into report"
            'Me.dgvMaster.Columns.Item("BOOLAPP").ToolTipText = "App: Set this as an appendix"
            'Me.dgvMaster.Columns.Item("BOOLFIG").ToolTipText = "Fig: Set this as a figure"
            'Me.dgvMaster.Columns.Item("BOOLIR").ToolTipText = "Incl: Include this Appendix/Figure in the report"
            'Me.dgvMaster.Columns.Item("CHARPAGEORIENTATION").ToolTipText = "P/L: Orient Appendix/Figure (P=Portrait, L=Landscape)"
            'Me.dgvMaster.Columns.Item("NUMSCALE").ToolTipText = "Scale the image on the page"
            'Me.dgvMaster.Columns.Item("NUMCROPLEFT").ToolTipText = "CL: Crop from left side (in inches)"
            'Me.dgvMaster.Columns.Item("NUMCROPRIGHT").ToolTipText = "CR: Crop from right side (in inches)"
            'Me.dgvMaster.Columns.Item("NUMCROPTOP").ToolTipText = "CT: Crop from top (in inches)"
            'Me.dgvMaster.Columns.Item("NUMCROPBOTTOM").ToolTipText = "CB: Crop from bottom (in inches)"
        Catch ex As Exception

        End Try

    End Sub

    Function FooterText() As String

        FooterText = ""

        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String

        Dim strM As String = ""
        Dim strF As String
        Dim strF1 As String
        Dim strF2 As String
        Dim var1, var2, var3, var4, var5

        Dim boolQuit As Boolean = False

        Dim intIt As Short = 0

        Dim strD1 As String
        Dim strD2 As String
        Dim strD3 As String
        Dim strD4 As String
        Dim strD5 As String

        Dim intMax As Short = 199

        Do Until boolQuit

            intIt = intIt + 1

            Select Case intIt

                Case 1

                    strD1 = "StudyDoc ID: "
                    strD2 = "Created Date: "
                    strD3 = "Generated Date: "
                    strD4 = "Generated By: "
                    strD5 = "Document Owner: "

                Case 2

                    strD1 = "StudyDoc ID: "
                    strD2 = "Created Date: "
                    strD3 = "Gen Date: "
                    strD4 = "Gen By: "
                    strD5 = "Doc Owner: "

                Case 3

                    boolQuit = True
                    FooterText = Mid(FooterText, 1, intMax)

                    str1 = "Note that the watermark content lenght exceeds the maximum number of characters (" & intMax & ")."
                    str1 = str1 & ChrW(10) & "The watermark content will be concatenated to " & intMax & " characters."
                    MsgBox(str1, vbInformation, "Invalid content...")

                    GoTo end1

            End Select

            strM = ""

            If boolNoneOp Then
                boolQuit = True
            Else

                Dim dt As Date = Now
                Dim dt1 As Date
                Dim strDate As String

                Dim idUA As Int64
                Dim idPers As Int64
                Dim strUID As String

                strF = "ID_TBLFINALREPORT = " & gid1
                Dim rows() As DataRow = tblFinalReport.Select(strF)
                strUID = NZ(rows(0).Item("CHARUSERID"), "NA")
                idPers = NZ(rows(0).Item("ID_TBLPERSONNEL"), 0)
                idUA = NZ(rows(0).Item("ID_TBLUSERACCOUNTS"), 0)
                strF1 = "ID_TBLPERSONNEL = " & idPers
                Dim rows1() As DataRow = tblPersonnel.Select(strF1)
                Dim rows2() As DataRow = tblUserAccounts.Select(strF1)

                If boolDocIDOp Then

                    str4 = strD1 & gid1

                    If Len(strM) = 0 Then
                        strM = str4
                    Else
                        If boolTextOp Or boolFooterOp Then
                            strM = strM & ", " & str4
                        Else
                            strM = strM & ChrW(10) & str4
                        End If

                    End If

                End If

                If boolDTCreatedOp Then
                    dt1 = rows(0).Item("UPSIZE_TS")

                    str1 = Format(dt1, LTextDateFormat)
                    str2 = Replace(Format(dt1, "hh:mm:ss tt"), " ", ChrW(160), 1, -1, CompareMethod.Text)

                    str4 = strD2 & str1 & ChrW(160) & str2

                    If Len(strM) = 0 Then
                        strM = str4
                    Else
                        If boolTextOp Or boolFooterOp Then
                            strM = strM & ", " & str4
                        Else
                            strM = strM & ChrW(10) & str4
                        End If
                    End If

                End If

                If boolDTReportedOp Then
                    str1 = Format(dt, LTextDateFormat)
                    str2 = Replace(Format(dt, "hh:mm:ss tt"), " ", ChrW(160), 1, -1, CompareMethod.Text)

                    str4 = strD3 & str1 & ChrW(160) & str2

                    If Len(strM) = 0 Then
                        strM = str4
                    Else
                        If boolTextOp Or boolFooterOp Then
                            strM = strM & ", " & str4
                        Else
                            strM = strM & ChrW(10) & str4
                        End If
                    End If


                End If

                If boolDocGenOp Then

                    str4 = strD4 & Replace(gUserName, " ", ChrW(160), 1, -1, CompareMethod.Text) & " as " & gUserID

                    If Len(strM) = 0 Then
                        strM = str4
                    Else
                        If boolTextOp Or boolFooterOp Then
                            strM = strM & ", " & str4
                        Else
                            strM = strM & ChrW(10) & str4 '& " and a bunch of other text to test length"
                        End If
                    End If



                End If

                If boolDocOwnerOp Then

                    If rows.Length = 0 Then
                        str4 = "NA"
                    Else
                        str1 = NZ(rows1(0).Item("CHARFIRSTNAME"), "NFN")
                        str2 = NZ(rows1(0).Item("CHARMIDDLENAME"), "NMN")
                        str3 = NZ(rows1(0).Item("CHARLASTNAME"), "NLN")

                        If Len(str2) = 0 Then
                            str4 = str1 & ChrW(160) & str3
                        Else
                            str4 = str1 & ChrW(160) & str2 & ChrW(160) & str3
                        End If

                        str4 = str4 & " as " & strUID

                    End If

                    str4 = strD5 & str4

                    If Len(strM) = 0 Then
                        strM = str4
                    Else
                        If boolTextOp Or boolFooterOp Then
                            strM = strM & ", " & str4
                        Else
                            strM = strM & ChrW(10) & str4
                        End If
                    End If


                End If

                If Len(strM) <= intMax Or boolTextOp Then
                    'character length restriction 199
                    boolQuit = True

                End If

                FooterText = strM

            End If

end1:

        Loop


        
        ''debuging
        'MsgBox(Len(FooterText))

        'str1 = "Add some text for testing"
        'str2 = InputBox(str1, "Enter")
        'FooterText = FooterText & str2

        'MsgBox(Len(FooterText))

    End Function



End Class