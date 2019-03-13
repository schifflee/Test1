Option Compare Text

Public Class frmAppFigs

    Public cbxTOC As New DataGridViewComboBoxCell
    Public cbxAnalyte As New DataGridViewComboBoxCell
    Public boolAppFigFormLoad As Boolean = False
    Private boolHold As Boolean = False

    Private Sub frmAppFigs_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed


    End Sub

    Private Sub frmAppFigs_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim str1 As String
        Dim str2 As String
        Dim strF As String
        Dim bool As Boolean
        Dim boolA As Boolean
        Dim w, h
        Dim var1

        Call ControlDefaults(Me)

        Call DoubleBufferControl(Me, "dgv")

        w = My.Computer.Screen.WorkingArea.Width
        h = My.Computer.Screen.WorkingArea.Height

        Me.Left = 0
        Me.Top = 0

        Me.Width = w '* 0.9
        Me.Height = h ' * 0.9

        str1 = "Note: If it is desired to create an Appendix section only with no inserted graphics, simply leave the Path column blank."
        Me.lblNote1.Text = str1

        boolAppFigFormLoad = True

        Cursor.Current = Cursors.AppStarting

        str1 = NZ(wWStudyName, "[NO STUDY CHOSEN]")

        str2 = "LABIntegrity StudyDoc" & ChrW(8482) & " - Appendices and Figures for " & str1

        Me.Text = str2

        Call InitializeTOC()

        Call InitializeTab1()

        'initialize Appendix
        'Call AppendixInitialize()

        Call MasterInitialize()

        Call FillMasterHelperText()

        Me.dgvTOC.CurrentCell = Me.dgvTOC.Rows.Item(0).Cells("CHARTITLE")

        Dim dgv As DataGridView
        Dim int1 As Short

        dgv = Me.dgvTOC
        int1 = dgv("INTTAB", 0).Value
        Me.tab1.SelectedTab = Me.tab1.TabPages.Item(int1)

        str1 = "* Optional Field Code ID: Used to create Field Code for this item"
        str1 = str1 & ChrW(10) & "* W: Insert figures from Word document"
        str1 = str1 & ChrW(10) & "* A: Order to add to report"
        str1 = str1 & ChrW(10) & "* App: Configure as an APPENDIX"
        str1 = str1 & ChrW(10) & "* Fig: Configure as a FIGURE"
        str1 = str1 & ChrW(10) & "* Incl: Include in the report"
        str1 = str1 & ChrW(10) & "* Page Orientation:  P=Portrait, L=Landscape"
        str1 = str1 & Chr(10) & "* CL,etc: Crop in inches"
        str1 = str1 & ChrW(10) & "** Optional"
        Me.lblLegend.Text = str1

        'position legend
        Me.gbxlblMasterAppFigs1.Top = Me.dgvMaster.Top - Me.gbxlblMasterAppFigs1.Height - 10  'NDL - Added 10 to give a little space.
        Me.gbxlblMasterAppFigs1.Left = Me.dgvMaster.Left + (Me.dgvMaster.Width - Me.gbxlblMasterAppFigs1.Width)

        boolA = BOOLREPORTTABLECONFIGURATION
        If boolA = 0 Then
            bool = False
        Else
            bool = True
        End If

        Me.cmdEdit.Enabled = bool

        Call GetDisplayAttachments()

        Call LockMaster(True)
        frmAppFigs_ToolTipSet()
        dgvMaster.AutoResizeRows()

        ''remove tabs 0 - 4
        Dim Count1 As Short
        For Count1 = 4 To 0 Step -1
            Me.tab1.TabPages.Remove(Me.tab1.TabPages(Count1))
        Next

        Try
            Me.tab1.Appearance = TabAppearance.Buttons
            Me.tab1.ItemSize = New Size(0, 1)
        Catch ex As Exception
            var1 = ex.Message
        End Try

        Cursor.Current = Cursors.Default

        'Set Captions for Word attachments as Read-Only and Grayed-out
        Dim row As DataGridViewRow

        For Each row In dgvMaster.Rows
            With row.Cells("CHARTITLE")
                Try
                    If row.Cells("BOOLW").Value = True Then
                        .Style.BackColor = Color.LightGray
                        str1 = "For Word attachments (W* checked), Caption is ignored."
                        str1 = str1 & ChrW(10) & "Captions must be included in the referenced document to appear in Report."
                        .ToolTipText = str1 ' "For Word attachments, Captions must also be inserted into the attached document to appear in Report."
                    Else
                        .Style.BackColor = Color.White
                        .ToolTipText = ""
                    End If
                Catch ex As Exception

                End Try

            End With
        Next

        Call SetComboCell(Me.dgvMaster, "CHARPAGEORIENTATION")

        Call DoWatsonRunNumberCol()

        Me.dgvMaster.AutoResizeColumns()
        Me.dgvMaster.AutoResizeRows()


        boolAppFigFormLoad = False
    End Sub

    Sub GetDisplayAttachments()

        If gboolDisplayAttachments Then
            Me.chkDisplayAttachment.Checked = True
        Else
            Me.chkDisplayAttachment.Checked = False
        End If
    End Sub

    Sub PutDisplayAttachments()

        If Me.chkDisplayAttachment.Checked Then
            gboolDisplayAttachments = True
        Else
            gboolDisplayAttachments = False
        End If

    End Sub

    Sub InitializeTOC()

        Dim dgv As DataGridView
        Dim dv As System.Data.DataView
        Dim Count1 As Short
        Dim int1 As Short
        Dim str1 As String

        dgv = Me.dgvTOC
        dgv.RowHeadersVisible = False
        dgv.ColumnHeadersVisible = False
        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells 'NDL - allowed to resize
        dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        'NDL - set these to false in the Designer.  Don't want user change row/column sizes
        '        dgv.AllowUserToResizeRows = True
        '        dgv.AllowUserToResizeColumns = True

        dv = New DataView(tblConfigAppFigs)
        dv.RowFilter = "BOOLMISC = 0 OR BOOLMISC = 1"
        dv.RowFilter = "ID_TBLCONFIGAPPFIGS > 4 AND ID_TBLCONFIGAPPFIGS < 9"
        dv.Sort = "CHARTITLE"
        dv.AllowDelete = False
        dv.AllowEdit = False
        dv.AllowNew = False

        dgv.DataSource = dv

        For Count1 = 0 To dgv.Columns.Count - 1
            dgv.Columns.Item(Count1).Visible = False
        Next
        dgv.Columns.Item("CHARTITLE").Visible = True
        dgv.Columns.Item("CHARINITIALS").Visible = True
        dgv.Columns.Item("CHARTITLE").MinimumWidth = dgv.Width * 0.75
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgv.Columns("CHARTITLE").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        dgv.Columns("CHARINITIALS").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        dgv.AutoResizeColumns()
        dgv.AutoResizeRows()

        'fill in cbxTOC
        Me.cbxTOC.Items.Clear()
        For Count1 = 0 To dv.Count - 1
            str1 = NZ(dv(Count1).Item("CHARINITIALS"), "")
            If Len(str1) = 0 Then
            Else
                int1 = dv(Count1).Item("BOOLMISC")
                If int1 = 1 Then
                Else
                    Me.cbxTOC.Items.Add(str1)
                End If
            End If
        Next
        Me.cbxTOC.Sorted = True
        Me.cbxTOC.AutoComplete = True
        Me.cbxTOC.MaxDropDownItems = 20
        Me.cbxTOC.DisplayStyleForCurrentCellOnly = True
        'Me.cbxTOC.DropDownWidth = Me.cbxTOC.DropDownWidth * 1.25
        Me.cbxTOC.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton

        'Me.dgvMaster("Value", int1) = Me.cbxTOC

        'fill in cbxanalyte
        Me.cbxAnalyte.Items.Clear()
        Me.cbxAnalyte.Items.Add("")
        For Count1 = 1 To ctAnalytes
            Me.cbxAnalyte.Items.Add(arrAnalytes(1, Count1))
        Next
        Me.cbxAnalyte.Sorted = True
        Me.cbxAnalyte.AutoComplete = True
        Me.cbxAnalyte.MaxDropDownItems = 20
        Me.cbxAnalyte.DisplayStyleForCurrentCellOnly = True
        Me.cbxAnalyte.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton


    End Sub

    Sub InitializeTab1()
        Dim tp As TabPage

        For Each tp In Me.tab1.TabPages
            tp.Text = ""
        Next

    End Sub

    Private Sub dgvTOC_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvTOC.SelectionChanged
        Dim dgv As DataGridView
        Dim int1 As Short
        Dim int2 As Short

        'NDL: Commented out these lines - currently, the TOC is inactive for this screen.
        'dgv = Me.dgvTOC
        'int2 = dgv.CurrentRow.Index
        'int1 = dgv("INTTAB", int2).Value

        'Me.tab1.SelectedTab = Me.tab1.TabPages.Item(int1)

        'dgv.Focus()


    End Sub

    Private Sub tab1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tab1.Click
        Dim int1 As Short
        Dim int2 As Short
        Dim Count1 As Short
        Dim dgv As DataGridView

        dgv = Me.dgvTOC

        int1 = Me.tab1.SelectedIndex
        int2 = dgv.Rows.Count
        'select appropriate dgvTOC row
        'dgv.CurrentCell = dgv.Rows.Item(int1).Cells("CHARTITLE")
        Try
            dgv.CurrentCell = dgv.Rows.Item(int1).Cells("CHARTITLE")
        Catch ex As Exception

        End Try

    End Sub


    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Visible = False
    End Sub


    Private Sub cmdMaster_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMaster.Click
        Dim dgv As DataGridView
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim int4 As Short
        Dim Count1 As Short

        dgv = Me.dgvTOC
        int1 = dgv.Rows.Count
        int2 = 14
        For Count1 = 0 To int1 - 1
            int3 = dgv("ID_TBLCONFIGAPPFIGS", Count1).Value
            If int3 = int2 Then
                int4 = dgv("INTTAB", Count1).Value
                dgv.Focus()
                dgv.CurrentCell = dgv.Rows.Item(Count1).Cells("CHARTITLE")
                Me.tab1.SelectedIndex = int4
                Exit For
            End If
        Next
    End Sub

    Sub DoThis(ByVal cmd As String)

        Dim boolMaster As Boolean

        boolMaster = True
        Select Case cmd
            Case "Edit"
                Me.SetToEditMode()
                boolMaster = False
            Case "Save"
                Me.SetToNonEditMode()

                strRFC = GetDefaultRFC()
                strMOS = GetDefaultMOS()

                Call SaveMaster()

                boolMaster = True
            Case "Cancel"
                Me.SetToNonEditMode()
                Call CancelMaster()

                boolMaster = True

        End Select

        Call LockMaster(boolMaster)

        Call DisplayIndex()

    End Sub

    Sub CancelMaster()

        tblAppFigs.RejectChanges()

        boolAppFigFormLoad = True
        Call MasterConfig()
        boolAppFigFormLoad = False

        'Me.dgvMaster.AutoResizeRows()


    End Sub

    Sub SaveMaster()

        Dim boolDo As Boolean = False
        Dim rows1() As DataRow = tblAppFigs.Select("", "", DataViewRowState.ModifiedOriginal)
        If rows1.Length = 0 Then
        Else
            boolDo = True
        End If

        'clear audittrailtemp
        tblAuditTrailTemp.Clear()
        idSE = 0

        Dim dgv As DataGridView

        dgv = Me.dgvMaster

        dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)

        Call FillAuditTrailTemp(tblAppFigs)

        If boolGuWuOracle Then
            Try
                ta_tblAppFigs.Update(tblAppFigs)
            Catch ex As DBConcurrencyException
                ''ds2005.TBLAPPFIGS.Merge(''ds2005.TBLAPPFIGS, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_tblAppFigsAcc.Update(tblAppFigs)
            Catch ex As DBConcurrencyException
                ''ds2005Acc.TBLAPPFIGS.Merge(''ds2005Acc.TBLAPPFIGS, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_tblAppFigsSQLServer.Update(tblAppFigs)
            Catch ex As DBConcurrencyException
                ''ds2005Acc.TBLAPPFIGS.Merge(''ds2005Acc.TBLAPPFIGS, True)
            End Try
        End If

        'record tblaudittrailtemp
        Call RecordAuditTrail(False, Now)


        '***tblReports

        Call FillAuditTrailTemp(tblReports)

        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String

        dtbl = tblReports
        strF = "ID_TBLSTUDIES = " & id_tblStudies
        rows = dtbl.Select(strF)

        If rows.Length = 0 Then

        Else
            rows(0).BeginEdit()
            If gboolDisplayAttachments Then
                rows(0).Item("BOOLDISPLAYATTACHMENTS") = -1
            Else
                rows(0).Item("BOOLDISPLAYATTACHMENTS") = 0
            End If
            rows(0).EndEdit()
        End If

        Dim dvCheck As System.Data.DataView = New DataView(tblReports)
        dvCheck.RowStateFilter = DataViewRowState.ModifiedCurrent
        Dim int10 As Short
        int10 = 1
        If int10 = 0 Then
        Else

            Call FillAuditTrailTemp(tblReports)

            If boolGuWuOracle Then
                Try
                    ta_tblReports.Update(tblReports)
                Catch ex As DBConcurrencyException
                    ''msgbox("aaHome: " & ex.Message)
                    ''ds2005.TBLREPORTS.Merge(''ds2005.TBLREPORTS, True)
                End Try

                Try
                    'ta_tblReportHeaders.Update(tblReportHeaders)
                Catch ex As DBConcurrencyException
                    ''ds2005.TBLREPORTHEADERS.Merge(''ds2005.TBLREPORTHEADERS, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_tblReportsAcc.Update(tblReports)
                Catch ex As DBConcurrencyException
                    ''msgbox("aaHome: " & ex.Message)
                    ''ds2005Acc.TBLREPORTS.Merge(''ds2005Acc.TBLREPORTS, True)
                End Try

                Try
                    ta_tblReportHeadersAcc.Update(tblReportHeaders)
                Catch ex As DBConcurrencyException
                    ''ds2005Acc.TBLREPORTHEADERS.Merge(''ds2005Acc.TBLREPORTHEADERS, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblReportsSQLServer.Update(tblReports)
                Catch ex As DBConcurrencyException
                    ''msgbox("aaHome: " & ex.Message)
                    ''ds2005Acc.TBLREPORTS.Merge(''ds2005Acc.TBLREPORTS, True)
                End Try

                Try
                    ta_tblReportHeadersSQLServer.Update(tblReportHeaders)
                Catch ex As DBConcurrencyException
                    ''ds2005Acc.TBLREPORTHEADERS.Merge(''ds2005Acc.TBLREPORTHEADERS, True)
                End Try
            End If

        End If

        'record tblaudittrailtemp
        Call RecordAuditTrail(False, Now)

        '***end tblReports

        If boolDo Then
            Call ResetFieldCodes(True)
        End If

        dgv.AutoResizeRows()


    End Sub

    Sub LockMaster(ByVal bool As Boolean)

        Me.cmdMasterInsert.Enabled = Not (bool)
        Me.cmdMasterDelete.Enabled = Not (bool)
        Me.cmdMasterBrowse.Enabled = Not (bool)
        Me.cmdResetMaster.Enabled = Not (bool)
        Me.cmdChrom.Enabled = Not (bool)

        Me.chkDisplayAttachment.Enabled = Not (bool)

        Me.dgvMaster.ReadOnly = bool
        Me.dgvMaster.Columns.Item("CHARPATH").ReadOnly = True


    End Sub

    Sub MasterInsertRow()

        Dim int1 As Short
        Dim int2 As Short
        Dim intNRNumber As Short
        Dim dv As System.Data.DataView
        Dim intOrder As Short
        Dim Count1 As Short
        Dim ct1 As Short
        Dim drows() As DataRow
        Dim strS As String

        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim maxID, maxID1

        Dim dgv As DataGridView
        Dim strTable As String
        Dim dtbl As System.Data.DataTable
        Dim strID As String
        Dim strID1 As String

        dgv = Me.dgvMaster
        strTable = "tblAppFigs"
        dtbl = tblAppFigs
        strID = "ID_TBLAPPFIGS"
        strID1 = "ID_TBLCONFIGAPPFIGS"

        maxID = 1
        maxID = GetMaxID(strTable, 1, True) 'if maxid increment is 1, then getmaxid already does putmaxid
        'Call PutMaxID(strTable, maxID)

        'If Len(strTable) = 0 Then
        'Else


        '    ''****
        '    'If boolGuWuOracle Then
        '    '    ta_tblMaxID.Fill(tblMaxID)
        '    'ElseIf boolGuWuAccess Then
        '    '    ta_tblMaxIDAcc.Fill(tblMaxID)
        '    'ElseIf boolGuWuSQLServer Then
        '    '    ta_tblMaxIDSQLServer.Fill(tblMaxID)
        '    'End If
        '    'strF = "charTable = '" & strTable & "'"
        '    'tbl = tblMaxID
        '    'rows = tbl.Select(strF)
        '    'maxID = rows(0).Item("NUMMAXID")
        '    'maxID1 = maxID
        '    'maxID = maxID + 1
        '    'rows(0).BeginEdit()
        '    'rows(0).Item("NUMMAXID") = maxID
        '    'rows(0).EndEdit()
        '    'If boolGuWuOracle Then
        '    '    Try
        '    '        ta_tblMaxID.Update(tblMaxID)
        '    '    Catch ex As DBConcurrencyException
        '    '        'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
        '    '    End Try
        '    'ElseIf boolGuWuAccess Then
        '    '    Try
        '    '        ta_tblMaxIDAcc.Update(tblMaxID)
        '    '    Catch ex As DBConcurrencyException
        '    '        'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
        '    '    End Try
        '    'ElseIf boolGuWuSQLServer Then
        '    '    Try
        '    '        ta_tblMaxIDSQLServer.Update(tblMaxID)
        '    '    Catch ex As DBConcurrencyException
        '    '        'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
        '    '    End Try
        '    'End If

        'End If

        dv = dgv.DataSource
        ct1 = dv.Count ' tbl.Rows.Count

        If dgv.SelectedRows.Count = 0 And ct1 = 0 Then
            int1 = 0
            intOrder = 1
        ElseIf dgv.SelectedRows.Count = 0 And ct1 > 0 Then
            int1 = ct1
            intOrder = ct1 + 1
        Else
            int1 = dgv.CurrentRow.Index + 1
            intOrder = dgv.CurrentRow.Index + 2
        End If

        strF = "id_tblStudies = " & id_tblStudies
        drows = dtbl.Select(strF, "intOrder ASC")
        'renumber rows
        For Count1 = int1 To ct1 - 1
            drows(Count1).BeginEdit()
            int2 = drows(Count1).Item("intOrder")
            drows(Count1).Item("intOrder") = int2 + 1
            drows(Count1).EndEdit()
        Next
        'now add row
        Dim row As DataRow = dtbl.NewRow()
        row.BeginEdit()
        If Len(strTable) = 0 Then
        Else
            row.Item(strID) = maxID
        End If
        row.Item("id_tblStudies") = id_tblStudies
        row.Item("intOrder") = intOrder
        row.Item("ID_TBLCONFIGAPPFIGS") = 0
        row.Item("CHARTYPE") = "MA"

        'populate default cell values
        row.Item("numScale") = 100
        row.Item("numCropLeft") = 0
        row.Item("numCropRight") = 0
        row.Item("numCropTop") = 0
        row.Item("numCropBottom") = 0
        row.Item("CHARPAGEORIENTATION") = "P"
        row.Item("BOOLINCLUDEINREPORT") = -1
        row.Item("BOOLIR") = True
        row.Item("BOOLAPP") = True
        row.Item("BOOLAPPENDIX") = -1
        row.Item("BOOLFIG") = False
        row.Item("BOOLFIGURE") = 0
        row.Item("BOOLINSERTWORDDOCS") = 0
        row.Item("BOOLW") = False
        row.EndEdit()
        dtbl.Rows.Add(row)

        dv = dtbl.DefaultView
        strF = "id_tblStudies = " & id_tblStudies
        strS = "intOrder ASC"
        dv.RowFilter = strF
        dv.Sort = strS
        dv.AllowNew = False
        dv.AllowEdit = True
        dv.AllowDelete = False
        'dgv.DataSource = dv
        Try
            dgv.DataSource = dv
            Call SetComboCell(dgv, "CHARPAGEORIENTATION")
        Catch ex As Exception

        End Try

    End Sub



    Sub MasterInitialize()

        Dim dgv As DataGridView


        dgv = Me.dgvMaster
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgv.AllowUserToResizeColumns = True
        dgv.AllowUserToResizeRows = True
        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        dgv.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True
        dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgv.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True

        dgv.AllowUserToResizeRows = True

        'dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

        dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        dgv.RowHeadersWidth = 20


        Call MasterConfig()


        Call MasterColFormat()

    End Sub

    Sub DisplayIndex()

        Dim dgv As DataGridView
        Dim str1 As String
        Dim Count1 As Short


        dgv = Me.dgvMaster

        Dim int1 As Short
        For Count1 = 1 To 3
            'this seems to be needed to be iterated
            int1 = 0

            int1 = int1 + 1
            dgv.Columns.Item("CHARFCID").DisplayIndex = int1
            int1 = int1 + 1
            dgv.Columns.Item("BOOLW").DisplayIndex = int1
            int1 = int1 + 1
            dgv.Columns.Item("CHARANALYTE").DisplayIndex = int1
            int1 = int1 + 1
            dgv.Columns.Item("CHARLMNUMBER").DisplayIndex = int1
            int1 = int1 + 1
            dgv.Columns.Item("NUMWATSONRUNNUMBER").DisplayIndex = int1
            int1 = int1 + 1
            dgv.Columns.Item("CHARTITLE").DisplayIndex = int1
            int1 = int1 + 1
            dgv.Columns.Item("CHARPATH").DisplayIndex = int1
            int1 = int1 + 1
            dgv.Columns.Item("INTORDER").DisplayIndex = int1
            int1 = int1 + 1
            dgv.Columns.Item("BOOLAPP").DisplayIndex = int1
            int1 = int1 + 1
            dgv.Columns.Item("BOOLFIG").DisplayIndex = int1
            int1 = int1 + 1
            dgv.Columns.Item("BOOLIR").DisplayIndex = int1
            int1 = int1 + 1
            dgv.Columns.Item("CHARPAGEORIENTATION").DisplayIndex = int1
            int1 = int1 + 1
            dgv.Columns.Item("NUMSCALE").DisplayIndex = int1
            int1 = int1 + 1
            dgv.Columns.Item("NUMCROPLEFT").DisplayIndex = int1
            int1 = int1 + 1
            dgv.Columns.Item("NUMCROPRIGHT").DisplayIndex = int1
            int1 = int1 + 1
            dgv.Columns.Item("NUMCROPTOP").DisplayIndex = int1
            int1 = int1 + 1
            dgv.Columns.Item("NUMCROPBOTTOM").DisplayIndex = int1
            int1 = int1 + 1

        Next

    End Sub

    Sub MasterColFormat()

        Dim dgv As DataGridView
        Dim str1 As String
        Dim Count1 As Short


        dgv = Me.dgvMaster
        dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True

        For Count1 = 0 To dgv.ColumnCount - 1
            dgv.Columns(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        Dim int1 As Short

        Call DisplayIndex()

        dgv.Columns.Item("ID_TBLAPPFIGS").Visible = False

        dgv.Columns.Item("ID_TBLCONFIGAPPFIGS").Visible = False

        dgv.Columns.Item("ID_TBLSTUDIES").Visible = False

        dgv.Columns.Item("NUMWATSONRUNNUMBER").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv.Columns.Item("NUMWATSONRUNNUMBER").Width = 60

        dgv.Columns.Item("CHARTYPE").Visible = True
        dgv.Columns.Item("CHARTYPE").DisplayIndex = 0
        dgv.Columns.Item("CHARTYPE").HeaderText = "Type"
        dgv.Columns.Item("CHARTYPE").Width = 75
        dgv.Columns.Item("INTORDER").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        dgv.Columns.Item("CHARANALYTE").Visible = False 'True
        dgv.Columns.Item("CHARANALYTE").HeaderText = "Analyte (Optional)"

        dgv.Columns.Item("CHARLMNUMBER").Visible = False
        dgv.Columns.Item("CHARLMNUMBER").HeaderText = "LM#"

        dgv.Columns.Item("NUMWATSONRUNNUMBER").Visible = False
        str1 = "Watson" & ChrW(13) & "Run ID" & ChrW(10) & "(Optional)"
        dgv.Columns.Item("NUMWATSONRUNNUMBER").HeaderText = str1

        dgv.Columns.Item("CHARTITLE").Visible = True
        dgv.Columns.Item("CHARTITLE").HeaderText = "Caption (Optional)"
        dgv.Columns.Item("CHARTITLE").MinimumWidth = 150
        dgv.Columns.Item("CHARTITLE").DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dgv.Columns.Item("CHARTITLE").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        dgv.Columns.Item("CHARPATH").Visible = True
        dgv.Columns.Item("CHARPATH").HeaderText = "Path (Optional) - Drag & Drop from Windows Explorer or Double-click on Cell"
        dgv.Columns.Item("CHARPATH").MinimumWidth = 250
        dgv.Columns.Item("CHARPATH").ReadOnly = True
        dgv.Columns.Item("CHARPATH").DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dgv.Columns.Item("CHARPATH").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        dgv.Columns.Item("INTORDER").Visible = True
        dgv.Columns.Item("INTORDER").HeaderText = "A*"
        dgv.Columns.Item("INTORDER").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        dgv.Columns.Item("INTORDER").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'dgv.Columns.item("INTORDER").Width = 35

        dgv.Columns.Item("BOOLAPP").Visible = True
        dgv.Columns.Item("BOOLAPP").HeaderText = "App*"
        dgv.Columns.Item("BOOLAPP").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        'dgv.Columns.item("BOOLAPP").Width = 35

        dgv.Columns.Item("BOOLFIG").Visible = True
        dgv.Columns.Item("BOOLFIG").HeaderText = "Fig*"
        dgv.Columns.Item("BOOLFIG").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        'dgv.Columns.item("BOOLFIG").Width = 35

        dgv.Columns.Item("BOOLIR").Visible = True
        dgv.Columns.Item("BOOLIR").HeaderText = "Incl.*"
        dgv.Columns.Item("BOOLIR").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        'dgv.Columns.item("BOOLIR").Width = 35

        dgv.Columns.Item("CHARFCID").Visible = True
        dgv.Columns.Item("CHARFCID").HeaderText = "Field Code" & ChrW(10) & "ID*"

        dgv.Columns.Item("BOOLW").Visible = True
        dgv.Columns.Item("BOOLW").HeaderText = "W*"
        dgv.Columns.Item("BOOLW").AutoSizeMode = DataGridViewAutoSizeColumnMode.None

        dgv.Columns.Item("BOOLAPPENDIX").Visible = False

        dgv.Columns.Item("BOOLFIGURE").Visible = False

        dgv.Columns.Item("BOOLINCLUDEINREPORT").Visible = False

        dgv.Columns.Item("BOOLINSERTWORDDOCS").Visible = False

        dgv.Columns.Item("CHARPAGEORIENTATION").Visible = True
        dgv.Columns.Item("CHARPAGEORIENTATION").HeaderText = "P/L*"
        dgv.Columns.Item("CHARPAGEORIENTATION").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        dgv.Columns.Item("CHARPAGEORIENTATION").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'dgv.Columns.item("CHARPAGEORIENTATION").Width = 35

        dgv.Columns.Item("NUMSCALE").Visible = True
        str1 = "Scale" & ChrW(13) & "(%)"
        str1 = "Scale (%)"
        dgv.Columns.Item("NUMSCALE").HeaderText = str1
        ' dgv.Columns.item("NUMSCALE").FillWeight = 90
        dgv.Columns.Item("NUMSCALE").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        dgv.Columns.Item("NUMSCALE").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'dgv.Columns.item("NUMSCALE").Width = 45

        dgv.Columns.Item("NUMCROPLEFT").Visible = True
        str1 = "CL" & ChrW(13) & "(in)"
        str1 = "CL (in)"
        dgv.Columns.Item("NUMCROPLEFT").HeaderText = str1
        'dgv.Columns.item("NUMCROPLEFT").FillWeight = 90
        dgv.Columns.Item("NUMCROPLEFT").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        dgv.Columns.Item("NUMCROPLEFT").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'dgv.Columns.item("NUMCROPLEFT").Width = 35

        dgv.Columns.Item("NUMCROPRIGHT").Visible = True
        str1 = "CR" & ChrW(13) & "(in)"
        str1 = "CR (in)"
        dgv.Columns.Item("NUMCROPRIGHT").HeaderText = str1
        'dgv.Columns.item("NUMCROPRIGHT").FillWeight = 90
        dgv.Columns.Item("NUMCROPRIGHT").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv.Columns.Item("NUMCROPRIGHT").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        'dgv.Columns.item("NUMCROPRIGHT").Width = 35

        dgv.Columns.Item("NUMCROPTOP").Visible = True
        str1 = "CT" & ChrW(13) & "(in)"
        str1 = "CT (in)"
        dgv.Columns.Item("NUMCROPTOP").HeaderText = str1
        'dgv.Columns.item("NUMCROPTOP").FillWeight = 90
        dgv.Columns.Item("NUMCROPTOP").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv.Columns.Item("NUMCROPTOP").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        'dgv.Columns.item("NUMCROPTOP").Width = 35

        dgv.Columns.Item("NUMCROPBOTTOM").Visible = True
        str1 = "CB" & ChrW(13) & "(in)"
        str1 = "CB (in)"
        dgv.Columns.Item("NUMCROPBOTTOM").HeaderText = str1
        'dgv.Columns.item("NUMCROPBOTTOM").FillWeight = 90
        dgv.Columns.Item("NUMCROPBOTTOM").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv.Columns.Item("NUMCROPBOTTOM").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        'dgv.Columns.item("NUMCROPBOTTOM").Width = 35


        dgv.Columns.Item("CHARRDB").Visible = True
        'dgv.Columns.Item("CHARRDB").HeaderText = "Sciex" & ChrW(8482) & " Analyst .rdb file"
        dgv.Columns.Item("CHARRDB").HeaderText = "Chromatogram Source File"
        dgv.Columns.Item("CHARRDB").MinimumWidth = 250
        dgv.Columns.Item("CHARRDB").ReadOnly = True
        dgv.Columns.Item("CHARRDB").DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dgv.Columns.Item("CHARRDB").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft


        dgv.Columns.Item("UPSIZE_TS").Visible = False

        'dgv.Columns.item("CHARTYPE").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        'dgv.Columns.item("CHARTITLE").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        'dgv.Columns.item("CHARTYPE").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

        dgv.AutoResizeColumns()
        dgv.AutoResizeRows()

    End Sub

    Sub MasterConfig()

        Dim dgv As DataGridView
        Dim strF As String
        Dim strS As String
        Dim tbl As System.Data.DataTable
        Dim str1 As String

        tbl = tblAppFigs
        'add unbound columns
        str1 = "BOOLAPP"
        If tbl.Columns.Contains(str1) Then
        Else
            Dim col As New DataColumn
            col.ColumnName = str1
            col.DataType = System.Type.GetType("System.Boolean")
            col.AllowDBNull = True 'False
            tbl.Columns.Add(col)

            str1 = "BOOLFIG"
            Dim col1 As New DataColumn
            col1.ColumnName = str1
            col1.DataType = System.Type.GetType("System.Boolean")
            col1.AllowDBNull = True 'False
            tbl.Columns.Add(col1)

            str1 = "BOOLIR" 'BOOLINCLUDEINREPORT
            Dim col2 As New DataColumn
            col2.ColumnName = str1
            col2.DataType = System.Type.GetType("System.Boolean")
            col2.AllowDBNull = True 'False
            tbl.Columns.Add(col2)

            str1 = "BOOLW" 'BOOLINCLUDEINREPORT
            Dim col3 As New DataColumn
            col3.ColumnName = str1
            col3.DataType = System.Type.GetType("System.Boolean")
            col3.AllowDBNull = True 'False
            tbl.Columns.Add(col3)


        End If

        Call FillMasterUnboundValues()

        dgv = Me.dgvMaster
        strF = "ID_TBLSTUDIES = " & id_tblStudies
        'strS = "BOOLFIGURE ASC, BOOLAPPENDIX ASC, ID_TBLCONFIGAPPFIGS ASC, INTORDER ASC"
        strS = "INTORDER ASC"
        Dim dv As System.Data.DataView = New DataView(tblAppFigs, strF, strS, DataViewRowState.CurrentRows)
        dv.AllowDelete = False
        dv.AllowNew = False
        dgv.DataSource = dv

        Dim var1, var2
        var1 = dgv.Rows.Count 'debug
        var2 = dv.Count

        Call SetComboCell(dgv, "CHARPAGEORIENTATION")

        dgv.AutoResizeRows()


    End Sub

    Sub FillMasterUnboundValues()

        Dim dgv As DataGridView
        Dim dv As System.Data.DataView
        Dim Count1 As Short
        Dim Count2 As Short
        Dim int1 As Short
        Dim strF As String
        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim var1
        Dim str1 As String
        Dim str2 As String

        tbl = tblAppFigs
        strF = "ID_TBLSTUDIES = " & id_tblStudies
        rows = tbl.Select(strF)
        str1 = ""
        str2 = ""
        For Count1 = 0 To rows.Length - 1
            rows(Count1).BeginEdit()
            For Count2 = 1 To 4
                Select Case Count2
                    Case 1
                        str1 = "BOOLINCLUDEINREPORT"
                        str2 = "BOOLIR"
                    Case 2
                        str1 = "BOOLAPPENDIX"
                        str2 = "BOOLAPP"
                    Case 3
                        str1 = "BOOLFIGURE"
                        str2 = "BOOLFIG"
                    Case 4
                        str1 = "BOOLINSERTWORDDOCS"
                        str2 = "BOOLW"

                End Select
                var1 = rows(Count1).Item(str1)
                If var1 = 0 Then
                    rows(Count1).Item(str2) = False
                Else
                    rows(Count1).Item(str2) = True
                End If
            Next
            rows(Count1).EndEdit()
        Next

        tblAppFigs.AcceptChanges()

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click


        'remove tabs 0 - 4
        Dim Count1 As Short
        Dim int1 As Short
        For Count1 = 4 To 0 Step -1
            Me.tab1.TabPages.Remove(Me.tab1.TabPages(Count1))
        Next

    End Sub

    Sub FillMasterHelperText()

        Dim Count1 As Short
        Dim Count2 As Short
        Dim var1, var2
        Dim dtbl As System.Data.DataTable
        Dim boolHit As Boolean
        Dim dgv As DataGridView
        Dim int1 As Short
        Dim int2 As Short
        Dim dv As System.Data.DataView
        Dim str1 As String
        Dim str2 As String
        Dim dgvH As DataGridView
        Dim tbl1 As New System.Data.DataTable
        Dim col1 As New DataColumn

        col1.ColumnName = "Text"
        tbl1.Columns.Add(col1)

        dtbl = modGlobal.tblAnalytesHome
        dgvH = Me.dgvMasterHelperText
        dgvH.RowHeadersVisible = False
        dgvH.ColumnHeadersVisible = False
        dgvH.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dgvH.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvH.SelectionMode = DataGridViewSelectionMode.CellSelect
        'dgvH.ReadOnly = True

        'fill analytes from AnalyteHome 
        For Count1 = 0 To dtbl.Rows.Count - 1
            var1 = CStr(dtbl.Rows.Item(Count1).Item("AnalyteDescription"))
            boolHit = False
            For Count2 = 0 To tbl1.Rows.Count - 1
                'var2 = CStr(NZ(dgvH("Text", Count2).Value, ""))
                var2 = NZ(tbl1.Rows.Item(Count2).Item("Text"), "")
                If StrComp(var1, var2, CompareMethod.Text) = 0 Then
                    boolHit = True
                    Exit For
                End If
            Next
            If boolHit Then 'ignore
            Else
                Dim row As DataRow = tbl1.NewRow
                row("Text") = var1
                tbl1.Rows.Add(row)
            End If
        Next

        'fill analytes from tblAnalref 
        dgv = frmH.dgvCompanyAnalRef
        dv = dgv.DataSource
        str1 = "Analyte Name"
        int1 = FindRowDVByCol(str1, dv, "Item")
        For Count1 = 0 To dgv.Columns.Count - 1
            str2 = dgv.Columns.Item(Count1).Name
            If StrComp(str2, "BOOLINCLUDE", CompareMethod.Text) = 0 Then 'IGNORE
            ElseIf StrComp(str2, "ID_TBLDATATABLEROWTITLES", CompareMethod.Text) = 0 Then 'IGNORE
            ElseIf StrComp(str2, "Item", CompareMethod.Text) = 0 Then 'IGNORE
            Else
                var1 = dgv(Count1, int1).Value
                boolHit = False
                For Count2 = 0 To tbl1.Rows.Count - 1
                    'var2 = CStr(NZ(dgvH("Text", Count2).Value, ""))
                    var2 = NZ(tbl1.Rows.Item(Count2).Item("Text"), "")
                    If StrComp(var1, var2, CompareMethod.Text) = 0 Then
                        boolHit = True
                        Exit For
                    End If
                Next
                If boolHit Then 'ignore
                Else
                    Dim row As DataRow = tbl1.NewRow
                    row("Text") = var1
                    tbl1.Rows.Add(row)
                End If
            End If
        Next

        'fill LM from methodval
        dtbl = tblMethodValData
        str1 = "Lab Method Number"
        int1 = FindRow(str1, dtbl, "Item")
        For Count1 = 1 To dtbl.Columns.Count - 1
            '20171110 LEE:
            'The following code throws an untrapped error if dtbl item is null
            'Amazing that an error was never thrown until now: Alturas
            'Need to add NZ()
            'var1 = CStr(dtbl.Rows.Item(int1).Item(Count1))
            Try
                var1 = CStr(NZ(dtbl.Rows.Item(int1).Item(Count1), ""))
            Catch ex As Exception
                var1 = var1
            End Try

            boolHit = False
            For Count2 = 0 To tbl1.Rows.Count - 1
                'var2 = CStr(NZ(dgvH("Text", Count2).Value, ""))
                var2 = NZ(tbl1.Rows.Item(Count2).Item("Text"), "")
                If StrComp(var1, var2, CompareMethod.Text) = 0 Then
                    boolHit = True
                    Exit For
                End If
            Next
            If boolHit Then 'ignore
            Else
                Dim row As DataRow = tbl1.NewRow
                row("Text") = var1
                tbl1.Rows.Add(row)
            End If
        Next

        Dim dvH As System.Data.DataView = New DataView(tbl1)
        Dim int3 As Short
        int3 = tbl1.Rows.Count
        dvH.AllowNew = False
        dvH.AllowDelete = False
        dvH.Sort = "Text ASC"
        dgvH.DataSource = dvH
        dgvH.AutoResizeRows()


    End Sub


    Private Sub dgvMaster_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvMaster.CellClick

        Dim str1 As String
        Dim str2 As String
        Dim dgv As DataGridView
        Dim boolGo As Boolean
        Dim var1

        If Me.cmdEdit.Enabled Or Len(frmH.cbxStudy.Text) = 0 Then
            Exit Sub
        End If

        If boolHold Then
            Exit Sub
        End If

        If e.RowIndex < 0 Then
            Exit Sub
        End If


        dgv = Me.dgvMaster
        'str1 = dgv.Columns.item(e.ColumnIndex).Name
        If e.ColumnIndex = -1 Then
            Exit Sub
        End If
        str1 = dgv.Columns.Item(e.ColumnIndex).Name

        If StrComp(str1, "CHARTYPE", CompareMethod.Text) = 0 Then

            str2 = NZ(dgv.Rows.Item(e.RowIndex).Cells(e.ColumnIndex).EditType.FullName, "")
            If InStr(1, str2, "combobox", CompareMethod.Text) > 0 Then
            Else
                Dim cbx As New DataGridViewComboBoxCell
                boolGo = False
                cbx = Me.cbxTOC.Clone
                'if data doesn't exist in dropdown list
                'data error will be called that inserts unlisted value into dropdown box
                On Error Resume Next
                dgv(e.ColumnIndex, e.RowIndex) = cbx
                If Err.Number <> 0 Then
                    Err.Clear()
                End If
                On Error GoTo 0

            End If
        ElseIf StrComp(str1, "CHARANALYTE", CompareMethod.Text) = 0 Then
            str2 = NZ(dgv.Rows.Item(e.RowIndex).Cells(e.ColumnIndex).EditType.FullName, "")
            If InStr(1, str2, "combobox", CompareMethod.Text) > 0 Then
            Else
                Dim cbx1 As New DataGridViewComboBoxCell
                boolGo = False
                cbx1 = Me.cbxAnalyte.Clone
                'if data doesn't exist in dropdown list
                'data error will be called that inserts unlisted value into dropdown box
                On Error Resume Next
                dgv(e.ColumnIndex, e.RowIndex) = cbx1
                If Err.Number <> 0 Then
                    Err.Clear()
                End If
                On Error GoTo 0

            End If

        End If
    End Sub

    Private Sub dgvMaster_CellChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvMaster.CellValueChanged

        Dim str1 As String

        If (e.RowIndex > -1) Then
            If (StrComp(Me.dgvMaster.Columns(e.ColumnIndex).Name, "BOOLW") = 0) Then

                'Set Captions for Word attachments as Read-Only and Grayed-out
                With Me.dgvMaster.Rows(e.RowIndex).Cells("CHARTITLE")
                    Try
                        If (Me.dgvMaster.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = True) Then
                            .Style.BackColor = Color.LightGray
                            'str1 = "For Word attachments, Captions must also be inserted into  the attached document to appear in Report."
                            str1 = "For Word attachments (W* checked), Caption is ignored."
                            str1 = str1 & ChrW(10) & "Captions must be included in the referenced document to appear in Report."
                            .ToolTipText = str1
                        Else
                            .Style.BackColor = Color.White
                            .ToolTipText = ""

                        End If
                    Catch ex As Exception

                    End Try

                End With
            End If
        End If
    End Sub

    Private Sub cmdMasterInsert_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdMasterInsert.Click

        Dim dgv As DataGridView
        Dim int1 As Short
        Dim ct1 As Short
        Dim dv As System.Data.DataView
        Dim int2 As Short

        dgv = Me.dgvMaster
        dv = dgv.DataSource
        ct1 = dv.Count ' tbl.Rows.Count

        If dgv.SelectedRows.Count = 0 And ct1 = 0 Then
            int1 = 0
        ElseIf dgv.SelectedRows.Count = 0 And ct1 > 0 Then
            int1 = ct1
        Else
            int1 = dgv.CurrentRow.Index + 1
        End If

        'Call GenericDGVRowInsert(dgv, tblAppFigs, "tblAppFigs", "ID_TBLAPPFIGS")
        boolHold = True
        Call MasterInsertRow()
        boolHold = False

        'dgv.Update()
        dgvMaster.AutoResizeRows()

        'select new row
        Try
            int2 = 0
            Do Until dgv.Columns.Item(int2).Visible
                int2 = int2 + 1
            Loop
            dgv.CurrentCell = dgv(int2, int1)
            dgv.Rows.Item(int1).Selected = True
        Catch ex As Exception

        End Try
    End Sub

    Private Sub cmdMasterDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdMasterDelete.Click
        Dim int1 As Short
        Dim ct1 As Short
        Dim dgv As DataGridView
        Dim int2 As Short

        dgv = Me.dgvMaster

        int1 = dgv.CurrentRow.Index

        boolHold = True
        Call GenericDGVRowDelete(dgv)

        ct1 = dgv.Rows.Count
        'select row before deleted row
        int2 = 0
        If ct1 = 0 Then
        ElseIf int1 = 0 Then
            Do Until dgv.Columns.Item(int2).Visible
                int2 = int2 + 1
            Loop
            dgv.CurrentCell = dgv(int2, int1)
            dgv.Rows.Item(int1).Selected = True
        Else
            Do Until dgv.Columns.Item(int2).Visible
                int2 = int2 + 1
            Loop

            Try
                dgv.CurrentCell = dgv(int2, int1 - 1)
                'dgv.Rows.Item(int1).Selected = True
                If int1 > dgv.RowCount - 1 Then
                    int1 = dgv.RowCount - 1
                    dgv.Rows.Item(int1).Selected = True

                Else
                    dgv.Rows.Item(int1).Selected = True
                End If
            Catch ex As Exception

            End Try
        End If
        boolHold = False

    End Sub

    Sub BrowseForPath()

        Dim str2 As String
        Dim int1 As Short
        Dim strF As String
        Dim strPath As String
        Dim frm As New frmBrowseAsk
        Dim dgv As DataGridView
        Dim strPath1 As String
        Dim strM As String

        dgv = Me.dgvMaster

        If dgv.Rows.Count < 1 Then
            MsgBox("Please insert a row...", MsgBoxStyle.Information, "Insert a row...")
            Exit Sub
        ElseIf dgv.CurrentRow Is Nothing Then
            MsgBox("Please select a row...", MsgBoxStyle.Information, "Select a row...")
            Exit Sub
        End If
        int1 = dgv.CurrentRow.Index

        frm.ShowDialog()
        If frm.boolGo Then 'continue
        Else
            frm.Dispose()
            Exit Sub
        End If

        Dim strFilter As String
        Dim strFileName As String
        strFilter = "All files (*.*)|*.*"
        strFileName = "*.*"
        strPath1 = NZ(dgv("CHARPATH", int1).Value, "")
        If (System.IO.File.Exists(strPath1)) Then   'NDL: For Word attachments which have a file specified.
            strPath1 = System.IO.Path.GetDirectoryName(strPath1)
        End If
        Dim boolIWD As Boolean = False
        Dim var1

        var1 = dgv("BOOLW", int1).Value
        If var1 Then
            boolIWD = True
        Else
            boolIWD = False
        End If

        If frm.boolAddPath Then

            'str1 = "To configure a directory, select a file in that directory when prompted with the Open File Dialog box..."
            'MsgBox(str1, MsgBoxStyle.Information, "Directory Configuration Instructions...")

            'get default path
            Dim dtbl As System.Data.DataTable
            dtbl = tblConfiguration
            strF = "charConfigTitle = 'Figures Path'"
            Dim rows() As DataRow
            rows = dtbl.Select(strF)
            strPath = NZ(rows(0).Item("charConfigValue"), "")
            strFileName = "*.*"
            If boolIWD Then
                'strFilter = ".doc(x) files (*.doc*)|*.doc*"
                strFilter = ".doc(x) files (*.doc*)|*.doc*"

                strFileName = "*.doc*"
                If Len(strPath1) = 0 Then
                    str2 = ReturnDirectoryBrowse(True, strPath, strFilter, strFileName, True) 'TRUE returns file
                Else
                    str2 = ReturnDirectoryBrowse(True, strPath1, strFilter, strFileName, True) 'TRUE returns file
                End If
            Else
                If Len(strPath1) = 0 Then
                    str2 = ReturnDirectoryBrowse(False, strPath, strFilter, strFileName, True) 'TRUE returns directory
                Else
                    str2 = ReturnDirectoryBrowse(False, strPath1, strFilter, strFileName, True) 'TRUE returns directory
                End If
            End If


            If Len(str2) = 0 Then
            Else

                If boolIWD Then
                    Dim strChk As String
                    strChk = Mid(str2, Len(str2) - 6, Len(str2))
                    If InStr(1, strChk, "doc", CompareMethod.Text) > 0 Then
                    Else
                        strM = "Document must be a Word" & ChrW(8482) & " document"
                        MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
                        GoTo end1
                    End If
                End If

                int1 = dgv.CurrentRow.Index
                Dim dv As System.Data.DataView
                dv = dgv.DataSource
                dv(int1).BeginEdit()
                dv(int1).Item("charPath") = str2
                dv(int1).EndEdit()
                dgv.AutoResizeRows()
            End If
        Else
            int1 = dgv.CurrentRow.Index
            Dim dv As System.Data.DataView
            dv = dgv.DataSource
            dv(int1).BeginEdit()
            dv(int1).Item("CHARPATH") = ""
            dv(int1).Item("CHARRDB") = ""
            dv(int1).EndEdit()
            dgv.AutoResizeRows()

        End If

end1:

        frm.Dispose()

    End Sub

    Private Sub cmdMasterBrowse_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdMasterBrowse.Click

        Call BrowseForPath()

    End Sub

    Private Sub cmdResetMaster_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdResetMaster.Click
        Call CancelMaster()
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
    Private Sub dgvMaster_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgvMaster.CellValidating
        If boolAppFigFormLoad Then
            Exit Sub
        End If

        If boolHold Then
            Exit Sub
        End If

        If Me.cmdEdit.Enabled Then
            Exit Sub
        End If

        Dim dgv As DataGridView
        Dim str1 As String
        Dim boolNumeric As Boolean
        Dim boolInteger As Boolean
        Dim var1
        Dim boolErr As Boolean
        Dim strM As String
        Dim varVal
        Dim varV
        Dim boolPL As Boolean
        Dim boolApp As Boolean
        Dim boolFig As Boolean
        Dim boolIR As Boolean
        Dim boolFCID As Boolean
        Dim boolW As Boolean
        Dim dv As System.Data.DataView
        Dim strAF As String
        Dim Count1 As Short
        Dim intRow As Short

        dgv = Me.dgvMaster
        str1 = dgv.Columns.Item(e.ColumnIndex).Name
        boolNumeric = False
        boolInteger = False
        boolErr = False
        boolPL = False
        boolApp = False
        boolFig = False
        boolIR = False
        boolFCID = False
        boolW = False

        intRow = e.RowIndex

        Select Case str1
            Case "INTORDER"
                boolInteger = True
            Case "NUMSCALE"
                boolNumeric = True
            Case "NUMCROPLEFT"
                boolNumeric = True
            Case "NUMCROPRIGHT"
                boolNumeric = True
            Case "NUMCROPTOP"
                boolNumeric = True
            Case "NUMCROPBOTTOM"
                boolNumeric = True
            Case "CHARPAGEORIENTATION"
                boolPL = True
            Case "BOOLFIG"
                boolFig = True
            Case "BOOLAPP"
                boolApp = True
            Case "BOOLIR"
                boolIR = True
            Case "CHARFCID"
                boolFCID = True
            Case "BOOLW"
                boolW = True
        End Select
        'varVal = dgv(e.ColumnIndex, e.RowIndex).Value
        varVal = e.FormattedValue
        strM = ""

        If boolFCID Then

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
            'must be unique in table
            dv = dgv.DataSource
            For Count1 = 0 To dv.Count - 1
                If Count1 = intRow Then 'ignore
                Else
                    varV = NZ(dv(Count1).Item("CHARFCID"), "")
                    If StrComp(CStr(varVal), CStr(varV), CompareMethod.Text) = 0 Then
                        strM = "Entry must be unique to this table"
                        boolErr = True
                        GoTo err1
                    End If
                End If

            Next

        End If


        If boolNumeric Then 'evaluate
            If IsNumeric(varVal) Then
            Else
                boolErr = True
                strM = "Entry must be numeric."
                GoTo err1
            End If
        End If

        If boolInteger Then 'evaluate
            If IsNumeric(varVal) Then
                If IsInt(varVal) Then
                Else
                    boolErr = True
                    strM = "Entry must be integer."
                    GoTo err1
                End If
            Else
                boolErr = True
                strM = "Entry must be integer."
                GoTo err1
            End If
        End If

        If boolPL Then
            If Len(varVal) > 1 Then
                boolErr = True
                strM = "Entry must be P or L."
                GoTo err1
            Else
                var1 = Asc(varVal)
                If var1 = 80 Or var1 = 76 Or var1 = 108 Or var1 = 112 Then 'pass
                    'make value P or L
                    If var1 = 108 Then
                        dgv(e.ColumnIndex, e.RowIndex).Value = "L"
                    ElseIf var1 = 112 Then
                        dgv(e.ColumnIndex, e.RowIndex).Value = "P"
                    End If
                Else
                    boolErr = True
                    strM = "Entry must be P or L."
                    GoTo err1
                End If
            End If

        End If

        If boolApp Or boolFig Then
            dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
            var1 = dgv("CHARTYPE", e.RowIndex).Value
            varVal = dgv("BOOLFIG", e.RowIndex).Value
            If StrComp(var1, "RC", CompareMethod.Text) = 0 Or StrComp(var1, "LM", CompareMethod.Text) = 0 Then
                'If varVal Then 'disallow
                '    If StrComp(var1, "RC", CompareMethod.Text) = 0 Then
                '        str1 = "chromatograms"
                '    Else
                '        str1 = "laboratory methods"
                '    End If
                '    strM = "Dude, " & str1 & " must be configured as an appendix"
                '    boolErr = True
                '    dgv("BOOLFIG", e.RowIndex).Value = False
                '    dgv("BOOLAPP", e.RowIndex).Value = True
                '    dgv("BOOLFIGURE", e.RowIndex).Value = 0
                '    dgv("BOOLAPPENDIX", e.RowIndex).Value = -1
                '    dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
                'End If
            End If
        End If

        If boolIR Then
            dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
            varVal = dgv("BOOLIR", e.RowIndex).Value
            If varVal Then
                varVal = -1
            End If
            dgv("BOOLINCLUDEINREPORT", e.RowIndex).Value = varVal
            dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If

        If boolW Then
            dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
            varVal = dgv("BOOLW", e.RowIndex).Value
            If varVal Then
                varVal = -1
            End If
            dgv("BOOLINSERTWORDDOCS", e.RowIndex).Value = varVal
            dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If


err1:
        If boolErr Then
            e.Cancel = True
            If Len(strM) = 0 Then
            Else
                MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
            End If
            If boolFig Or boolApp Or boolW Then
            Else
                SendKeys.Send("{ESC}")
            End If


        End If

    End Sub

    Private Sub dgvMaster_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvMaster.MouseEnter
        Me.dgvMaster.Focus()
    End Sub

    Private Sub dgvMaster_RowValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvMaster.RowValidating

        'check to insure CHARTYPE is not null
        Dim dgv As DataGridView
        Dim var1
        Dim strM As String
        Dim boolErr As Boolean
        Dim str1 As String
        Dim int1 As Short
        Dim dv As System.Data.DataView
        Dim int2 As Long
        Dim intRow As Short
        Dim strSel As String

        If boolAppFigFormLoad Then
            Exit Sub
        End If

        If boolHold Then
            Exit Sub
        End If


        dgv = Me.dgvMaster
        If dgv.Rows.Count = 0 Then
            Exit Sub
        End If

        intRow = dgv.CurrentRow.Index
        var1 = NZ(dgv("CHARTYPE", intRow).Value, "")
        boolErr = False
        strM = ""
        strSel = ""
        If Len(var1) = 0 Then
            boolErr = True
            strM = "Type cannot be blank"
            strSel = "CHARTYPE"
            GoTo err1
        Else 'update ID_TBLCONFIGAPPFIGS	
            str1 = "ID_TBLCONFIGAPPFIGS"
            var1 = dgv("CHARTYPE", intRow).Value
            'find var1 in dgvTOC
            dv = Me.dgvTOC.DataSource
            int1 = FindRowDVByCol(var1, dv, "CHARINITIALS")
            'int2 = dv(int1).Item("ID_TBLCONFIGAPPFIGS")
            'dgv("ID_TBLCONFIGAPPFIGS", intRow).Value = int2
            'dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
            If int1 = -1 Then
            Else
                int2 = dv(int1).Item("ID_TBLCONFIGAPPFIGS")
                dgv("ID_TBLCONFIGAPPFIGS", intRow).Value = int2
                dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
            End If
        End If

        '20151027 Larry: Deprecate this feature
        'Caption can be blank for Word documents

        ''****
        'var1 = NZ(dgv("CHARTITLE", intRow).Value, "")
        'boolErr = False
        'strM = ""
        'If Len(var1) = 0 Then
        '    boolErr = True
        '    strM = "Caption cannot be blank"
        '    strSel = "CHARTITLE"
        '    GoTo err1
        'Else 'update ID_TBLCONFIGAPPFIGS	
        'End If

        ''*****
err1:
        If boolErr Then
            e.Cancel = True
            MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
            'dgv.CurrentCell = dgv.Rows.item(intRow).Cells(strSel)
            Try
                dgv.CurrentCell = dgv.Rows.Item(intRow).Cells(strSel)
            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        Dim var1
        var1 = NZ(frmH.cbxStudy.Text, "")
        If Len(var1) = 0 Then
            MsgBox("A study must be chosen.", MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If
        Call DoThis("Edit")
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Call DoThis("Save")
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        boolHold = True
        Call DoThis("Cancel")
        boolHold = False

    End Sub

    Private Sub dgvMaster_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvMaster.Click
        Dim dgv As DataGridView
        Dim intR As Short
        Dim intC As Short
        Dim var1, var2
        Dim str1 As String

        If boolHold Then
            Exit Sub
        End If

        dgv = Me.dgvMaster
        If dgv.Rows.Count = 0 Then
            Exit Sub
        ElseIf dgv.CurrentRow Is Nothing Then
            Exit Sub
        Else
            intR = dgv.CurrentRow.Index
            intC = dgv.CurrentCell.ColumnIndex
        End If

        var2 = NZ(dgv("CHARTYPE", intR).Value, "")
        If Len(var2) = 0 Then
        Else
            If StrComp(var2, "RC", CompareMethod.Text) = 0 Then
                'show repchromrun column
                'dgv.Columns.item("NUMWATSONRUNNUMBER").Visible = True
                'dgv("BOOLAPP", intR).Value = True
                'dgv("BOOLFIG", intR).Value = False
                'dgv("BOOLAPPENDIX", intR).Value = -1
                'dgv("BOOLFIGURE", intR).Value = 0
                'dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
                dgv.Columns.Item("NUMWATSONRUNNUMBER").Visible = True
            Else
                'dgv.Columns.item("NUMWATSONRUNNUMBER").Visible = False
                Try
                    dgv.Columns.Item("NUMWATSONRUNNUMBER").Visible = False
                Catch ex As Exception

                End Try
            End If

            'If StrComp(var2, "LM", CompareMethod.Text) = 0 Then
            '    'show repchromrun column
            '    dgv("BOOLAPP", intR).Value = True
            '    dgv("BOOLFIG", intR).Value = False
            '    dgv("BOOLAPPENDIX", intR).Value = -1
            '    dgv("BOOLFIGURE", intR).Value = 0
            '    dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
            'End If

        End If

    End Sub


    Private Sub dgvMaster_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvMaster.CellContentClick

        If boolAppFigFormLoad Then
            Exit Sub
        End If

        If boolHold Then
            Exit Sub
        End If

        If Me.cmdEdit.Enabled Then
            Exit Sub
        End If

        If e.RowIndex = -1 Then
            Exit Sub
        End If

        If e.ColumnIndex = -1 Then
            Exit Sub
        End If
        Dim dgv As DataGridView
        Dim var1, var2
        Dim str1 As String

        dgv = Me.dgvMaster
        str1 = dgv.Columns.Item(e.ColumnIndex).Name

        Select Case str1
            Case "BOOLAPP"

                dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
                var1 = dgv("BOOLAPP", e.RowIndex).Value

                'var2 = Not (var1)
                If var1 Then
                    'dgv("BOOLAPP", e.RowIndex).Value = var2
                    dgv("BOOLAPPENDIX", e.RowIndex).Value = -1
                    dgv("BOOLFIGURE", e.RowIndex).Value = 0
                    dgv("BOOLFIG", e.RowIndex).Value = Not (var1)
                Else
                    'dgv("BOOLAPP", e.RowIndex).Value = var1
                    dgv("BOOLAPPENDIX", e.RowIndex).Value = 0
                    dgv("BOOLFIGURE", e.RowIndex).Value = -1
                    dgv("BOOLFIG", e.RowIndex).Value = Not (var1)
                End If

            Case "BOOLFIG"

                dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
                var1 = dgv("BOOLFIG", e.RowIndex).Value

                'var2 = Not (var1)
                If var1 Then
                    dgv("BOOLAPP", e.RowIndex).Value = Not (var1)
                    dgv("BOOLAPPENDIX", e.RowIndex).Value = 0
                    dgv("BOOLFIGURE", e.RowIndex).Value = -1
                    'dgv("BOOLFIG", e.RowIndex).Value = var2
                Else
                    dgv("BOOLAPP", e.RowIndex).Value = Not (var1)
                    dgv("BOOLAPPENDIX", e.RowIndex).Value = -1
                    dgv("BOOLFIGURE", e.RowIndex).Value = 0
                    'dgv("BOOLFIG", e.RowIndex).Value = var1
                End If
        End Select

        dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)

    End Sub

    Private Sub dgvMasterHelperText_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvMasterHelperText.CellBeginEdit
        e.Cancel = True
    End Sub


    Private Sub dgvMaster_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvMaster.DataError
        'Dim var1
        'Dim var2
        'Dim dgv As DataGridView
        'Dim str1 As String
        'Dim boolGo As Boolean
        'Dim cbx As DataGridViewComboBoxCell
        'Dim cbx1 As New DataGridViewComboBoxCell

        'dgv = Me.dgvMaster
        'var1 = dgv.Rows.item(e.RowIndex).Cells(e.ColumnIndex).Value
        'str1 = dgv.Columns.item(e.ColumnIndex).Name
        'boolGo = False
        'Select Case str1
        '    Case "CHARTYPE"
        '        cbx = Me.cbxTOC
        '        boolGo = True
        'End Select
        'If boolGo Then
        '    cbx.Items.Add(var1)
        '    cbx1 = cbx.Clone
        '    dgv(e.ColumnIndex, e.RowIndex) = cbx1
        '    dgv(e.ColumnIndex, e.RowIndex).Value = var1
        'End If
    End Sub

    Private Sub cmiHomeFieldCode_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmiHomeFieldCode.Click
        Call InsertFC()
    End Sub

    Sub InsertFC()
        'Dim pos As Short
        'Dim strT As String
        'Dim str1 As String
        'Dim strL As String
        'Dim strR As String
        'Dim dgv As DataGridView
        'Dim intRow As Short
        'Dim intCol As Short


        'dgv = Me.dgvMaster
        'If dgv.RowCount = 0 Then
        '    Exit Sub
        'End If

        'If dgv.CurrentRow Is Nothing Then
        '    Exit Sub
        'End If

        'intRow = dgv.CurrentRow.Index
        'intCol = 3

        'strT = dgv.CurrentCell.ToString


        ''record position of cursor in text box
        'pos = Me.txtTitle.SelectionStart + Me.txtTitle.SelectionLength

        'Me.txtTitle.SelectionLength = 0
        'Me.txtTitle.SelectionStart = pos

        'Dim frm As New frmFieldCodes

        'Me.Cursor = New Cursor(Cursor.Current.Handle)

        ''frm.Location = new system.drawing.point(l1, t1)

        'frm.Location = new system.drawing.point(Cursor.Position.X, Cursor.Position.Y + 10)

        'frm.ShowDialog()

        'If frm.boolCancel Then

        '    Me.txtTitle.SelectionStart = pos
        '    Me.txtTitle.SelectionLength = 0

        'Else
        '    strT = Me.txtTitle.Text

        '    If pos = 0 Then
        '        strL = "" 'Mid(strT, 1, pos)
        '        strR = strT 'Mid(strT, pos + 1, Len(strT) - pos)
        '        str1 = frm.strFC & " " & strR
        '    ElseIf pos = Len(strT) Then
        '        strL = strT 'Mid(strT, 1, pos)
        '        strR = "" 'Mid(strT, pos + 1, Len(strT) - pos)
        '        str1 = strL & " " & frm.strFC
        '    Else
        '        strL = Mid(strT, 1, pos - 1)
        '        strR = Mid(strT, pos + 1, Len(strT) - pos)
        '        str1 = strL & " " & frm.strFC & " " & strR
        '    End If

        '    'strL = Mid(strT, 1, pos - 1)
        '    'strR = Mid(strT, pos + 1, Len(strT) - pos)
        '    str1 = strL & " " & frm.strFC & " " & strR
        '    Me.txtTitle.Text = str1

        'End If

        'frm.Dispose()
    End Sub

    Private Sub cmdClipboard_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClipboard.Click

        Call CopyToClipboard(Me.dgvMaster)

    End Sub

    Sub CopyToClipboard(ByVal dgv As DataGridView)

        Dim var1, var2, var3
        Dim strN As String
        Dim Count1 As Short
        Dim intRows As Short
        Dim strC As String
        Dim strT As String

        strN = dgv.Name
        strC = "Type" & ChrW(9) & "Field Code ID" & ChrW(9) & "Watson Run Id" & ChrW(9) & "App/Fig Description" & ChrW(9) & "Is Figure?" & ChrW(9) & "Is Appendix?"
        Select Case strN
            Case "dgvMaster"
                intRows = dgv.RowCount
                For Count1 = 0 To intRows - 1
                    var1 = dgv("CHARTYPE", Count1).Value
                    strT = CStr(NZ(var1, ""))
                    var1 = dgv("CHARFCID", Count1).Value
                    strT = strT & ChrW(9) & CStr(NZ(var1, ""))
                    var1 = dgv("NUMWATSONRUNNUMBER", Count1).Value
                    strT = strT & ChrW(9) & CStr(NZ(var1, ""))
                    var1 = dgv("CHARTITLE", Count1).Value
                    strT = strT & ChrW(9) & CStr(NZ(var1, ""))
                    var1 = dgv("BOOLFIGURE", Count1).Value
                    If var1 = 0 Then
                        var1 = "No"
                    Else
                        var1 = "Yes"
                    End If
                    strT = strT & ChrW(9) & CStr(NZ(var1, 0))
                    var1 = dgv("BOOLAPPENDIX", Count1).Value
                    If var1 = 0 Then
                        var1 = "No"
                    Else
                        var1 = "Yes"
                    End If
                    strT = strT & ChrW(9) & CStr(NZ(var1, 0))
                    strC = strC & ChrW(10) & strT
                Next
        End Select

        Try
            Clipboard.Clear()
        Catch ex As Exception

        End Try

        Try
            Clipboard.SetText(strC)
        Catch ex As Exception

        End Try

    End Sub

    Private Sub dgvMaster_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvMaster.CellContentDoubleClick

        If boolAppFigFormLoad Then
            Exit Sub
        End If

        If boolHold Then
            Exit Sub
        End If

        If Me.cmdEdit.Enabled Then
            Exit Sub
        End If

        If e.RowIndex = -1 Then
            Exit Sub
        End If

        If e.ColumnIndex = -1 Then
            Exit Sub
        End If
        Dim dgv As DataGridView
        Dim var1, var2
        Dim str1 As String

        dgv = Me.dgvMaster
        str1 = dgv.Columns.Item(e.ColumnIndex).Name

        Select Case str1
            Case "CHARPATH"

                Call BrowseForPath()

        End Select

    End Sub

    Private Sub chkDisplayAttachment_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDisplayAttachment.CheckedChanged

        Call PutDisplayAttachments()

    End Sub

    Private Sub dgvMaster_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvMaster.CellDoubleClick

        If Me.cmdEdit.Enabled Then
            Exit Sub
        End If

        If Me.cmdEdit.Enabled = False And Me.cmdSave.Enabled = False Then
            Exit Sub
        End If

        If e.RowIndex = -1 Then
            Exit Sub
        End If

        If e.ColumnIndex = -1 Then
            Exit Sub
        End If
        Dim dgv As DataGridView
        Dim str1 As String

        dgv = Me.dgvMaster
        str1 = dgv.Columns.Item(e.ColumnIndex).Name

        Select Case str1
            Case "CHARPATH"

                Call BrowseForPath()

        End Select

    End Sub

    ' Nick Addition
    Private Sub SetToEditMode()

        Me.cmdEdit.Enabled = False
        Me.cmdEdit.BackColor = System.Drawing.Color.Gray
        Me.cmdSave.Enabled = True
        Me.cmdSave.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCancel.Enabled = True
        Me.cmdCancel.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdExit.Enabled = False
        Me.cmdExit.BackColor = System.Drawing.Color.Gray
        Me.dgvMaster.AllowDrop = True

    End Sub

    Private Sub SetToNonEditMode()

        Me.cmdEdit.Enabled = True
        Me.cmdEdit.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdSave.Enabled = False
        Me.cmdSave.BackColor = System.Drawing.Color.Gray
        Me.cmdCancel.Enabled = False
        Me.cmdCancel.BackColor = System.Drawing.Color.Gray
        Me.cmdExit.Enabled = True
        Me.cmdExit.BackColor = System.Drawing.Color.Gainsboro
        Me.dgvMaster.AllowDrop = False
    End Sub

    Private Sub frmAppFigs_ToolTipSet()
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
            toolTip1.SetToolTip(Me.Label38, "Add Appendices and Figures (from other files) to the report")
            toolTip1.SetToolTip(Me.cmdClipboard, "Copy appendices/figures table to clipboard")
            toolTip1.SetToolTip(Me.cmdChrom, "Configure a chromatogram Word document obtained from ChromReporter" & ChrW(8482))
            toolTip1.SetToolTip(Me.cmdMasterInsert, "Insert appendix/figure below selected row")
            toolTip1.SetToolTip(Me.cmdMasterDelete, "Delete selected row")
            toolTip1.SetToolTip(Me.cmdMasterBrowse, "Find file for directory path in selected row")
            toolTip1.SetToolTip(Me.cmdEdit, "Change to Editing Mode")
            toolTip1.SetToolTip(Me.cmdSave, "Save all changes")
            toolTip1.SetToolTip(Me.cmdCancel, "Cancel Unsaved Changes")
            toolTip1.SetToolTip(Me.cmdExit, "Exit Appendices & Figures")
            toolTip1.SetToolTip(Me.cmdResetMaster, "Undo unsaved changes (this page only)")
            'Grid
            Me.dgvMaster.Columns.Item("CHARTYPE").ToolTipText = "Choose type of Appendix/Figure"
            Me.dgvMaster.Columns.Item("CHARFCID").ToolTipText = "Choose a field code ID for Word Template"
            Me.dgvMaster.Columns.Item("BOOLW").ToolTipText = "W: Insert figures from a word document"
            str1 = "Enter Caption for the Appendix/Figure."
            str1 = str1 & ChrW(10) & "Note that Caption is ignored if W* column is checked."
            Me.dgvMaster.Columns.Item("CHARTITLE").ToolTipText = str1 '"Enter Title for the Appendix/Figure"
            Me.dgvMaster.Columns.Item("CHARPATH").ToolTipText = "Choose directory path for supporting file"""
            Me.dgvMaster.Columns.Item("INTORDER").ToolTipText = "A: Order in which appendices/figures are put into report"
            Me.dgvMaster.Columns.Item("BOOLAPP").ToolTipText = "App: Set this as an appendix"
            Me.dgvMaster.Columns.Item("BOOLFIG").ToolTipText = "Fig: Set this as a figure"
            Me.dgvMaster.Columns.Item("BOOLIR").ToolTipText = "Incl: Include this Appendix/Figure in the report"
            Me.dgvMaster.Columns.Item("CHARPAGEORIENTATION").ToolTipText = "P/L: Orient Appendix/Figure (P=Portrait, L=Landscape)"
            Me.dgvMaster.Columns.Item("NUMSCALE").ToolTipText = "Scale the image on the page"
            Me.dgvMaster.Columns.Item("NUMCROPLEFT").ToolTipText = "CL: Crop from left side (in inches)"
            Me.dgvMaster.Columns.Item("NUMCROPRIGHT").ToolTipText = "CR: Crop from right side (in inches)"
            Me.dgvMaster.Columns.Item("NUMCROPTOP").ToolTipText = "CT: Crop from top (in inches)"
            Me.dgvMaster.Columns.Item("NUMCROPBOTTOM").ToolTipText = "CB: Crop from bottom (in inches)"
        Catch ex As Exception

        End Try

    End Sub


    Private Sub dgvMaster_DragDrop(sender As Object, e As DragEventArgs) Handles dgvMaster.DragDrop

        Dim path As String
        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
        Dim rows() As DataGridViewRow
        Dim intPathColumn, intWordColumn, intRow As Integer
        Dim pointHit As Point
        Dim strPath As String
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim str5 As String
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short

        int1 = 15
        int2 = 30
        int3 = 45

        If (cmdCancel.Enabled = False) Then  'Not in edit mode
            Exit Sub
        End If

        intPathColumn = Me.dgvMaster.Columns.Item("CHARPATH").Index
        intWordColumn = Me.dgvMaster.Columns.Item("BOOLW").Index
        pointHit = dgvMaster.PointToClient(New Point(e.X, e.Y))
        intRow = dgvMaster.HitTest(pointHit.X, pointHit.Y).RowIndex
        If intRow > -1 Then
            For Each path In files
                'Check if we are doing directories or paths
                strPath = path
                str1 = Mid(path, 1, int1) & "..."
                str2 = Mid(path, int1 + 1, int1) & ". ."
                str3 = Mid(path, int2 + 1, int1) & ". ."
                str4 = Mid(path, int3 + 1, int1) & ". ."
                str5 = Mid(path, int3 + 1 + int1, Len(path))
                strPath = str1 & str2 & str3 & str4 & str5 & "\"

                If (dgvMaster.Item(intWordColumn, intRow).Value) Then 'Its a Word Document
                    If (My.Computer.FileSystem.FileExists(path)) Then
                        Me.dgvMaster.Item(intPathColumn, intRow).Value = IBS(path) ' path
                    Else
                        MsgBox("Please drag a File (not a folder) for Word-based Figures & Appendices.")
                    End If
                Else
                    If (My.Computer.FileSystem.DirectoryExists(path)) Then
                        Me.dgvMaster.Item(intPathColumn, intRow).Value = IBS(path) 'path
                    Else
                        MsgBox("Please drag a Folder (not a file) for image-based Figures Appendices.")
                    End If
                End If
                Exit For
            Next
        End If

        'Me.dgvMaster.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders


        ''do again
        ''Me.dgvMaster.AutoResizeRows()

        'Dim dgv As DataGridView


        'dgv = Me.dgvMaster
        'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        'dgv.AllowUserToResizeColumns = True
        'dgv.AllowUserToResizeRows = True
        'dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        'dgv.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True
        'dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        'dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        'dgv.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True

        'dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

        'dgv.AllowUserToResizeRows = True

        'dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        'dgv.RowHeadersWidth = 20



    End Sub


    Private Sub dgvMaster_DragEnter(sender As Object, e As DragEventArgs) Handles dgvMaster.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub


    Private Sub cmdChrom_Click(sender As Object, e As EventArgs) Handles cmdChrom.Click

        Dim strCRPath As String = GetChromReporter()

        If Len(strCRPath) = 0 Then
            GoTo end1
        End If

        'now check to ensure the row is configured correctly

        Dim dgv As DataGridView = Me.dgvMaster
        Dim int1 As Int16
        Dim intRow As Int16
        Dim str1 As String
        Dim var1
        Dim strM As String

        strM = "Row must be configured as RC - Representative Chromatogram."

        Try

            intRow = dgv.CurrentRow.Index
            var1 = dgv("CHARTYPE", intRow).Value
            str1 = NZ(var1, "")
            If Len(str1) = 0 Then
                MsgBox(strM, vbInformation, "Invalid action...")
                GoTo end1
            End If

            If StrComp(str1, "RC", CompareMethod.Text) = 0 Then
            Else
                MsgBox(strM, vbInformation, "Invalid action...")
                GoTo end1
            End If

            Call DoChromReporter(strCRPath, intRow)

        Catch ex As Exception

        End Try



end1:

    End Sub

    Sub DoChromReporter(strCRPath As String, intRow As Int16)

        Dim frm As New frmChromReporterChoice

        frm.ShowDialog()

        If frm.boolCancel Then
            GoTo end1
        End If

        'enter information
        Dim dgv As DataGridView = Me.dgvMaster
        Dim dv As DataView = dgv.DataSource

        Dim strRDB As String = frm.txtRDB.Text
        Dim strPath As String = frm.txtDestinationPath.Text
        Dim strFileName As String = frm.txtWordFileName.Text
        Dim strFN As String
        Dim str1 As String
        Dim var1
        Dim strM As String

        frm.Dispose()

        Dim strFullPath As String

        Dim strDocx As String
        If Len(strFileName) < 5 Then
            strFN = strFileName & ".docx"
        Else
            str1 = Mid(strFileName, Len(strFileName) - 4, 5)
            If StrComp(str1, ".docx", CompareMethod.Text) = 0 Then
                strFN = strFileName
            Else
                strFN = strFileName & ".docx"
            End If
        End If

        strFullPath = EnterBackSlash(strPath) & strFN


        dv(intRow).Item("CHARPATH") = strFullPath
        dv(intRow).Item("CHARRDB") = strRDB

        'If (StrComp(Me.dgvMaster.Columns(e.ColumnIndex).Name, "BOOLW") = 0) Then
        Try
            dv(intRow).BeginEdit()
            dv(intRow).Item("BOOLW") = True
            dv(intRow).Item("BOOLINSERTWORDDOCS") = -1
            dv(intRow).EndEdit()
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try


        dgv.AutoResizeRows()
        dgv.AutoResizeColumns()

        'call ChromReporter
        Dim strPathCR As String = GetChromReporter()
        If System.IO.File.Exists(strPathCR) Then
        Else
            strM = "The ChromReport executable:" & ChrW(10) & ChrW(10) & strPathCR & ChrW(10) & ChrW(10) & "does not exist."
            MsgBox(strM, vbInformation, "Invalid action...")
            GoTo end1
        End If

        Dim proc As New System.Diagnostics.Process()
        'Public Shared Function Start(fileName As String, arguments As String) As Process

        Dim strArg1 As String 'rdb path
        Dim strArg2 As String 'Word path
        Dim strArg3 As String 'App
        Dim strArg As String

        strArg1 = "/RDB """ & strRDB & """"

        strArg2 = "/WORD """ & strFullPath & """"

        strArg3 = "/APP ""StudyDoc"""

        strArg = strArg1 & " " & strArg2 & " " & strArg3

        Try
            proc = Process.Start(strPathCR, strArg)
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try


end1:


    End Sub



    Sub DoWatsonRunNumberCol()

        Dim dgv As DataGridView = Me.dgvMaster
        Dim intRow As Int16
        Dim str1 As String

        Try
            intRow = dgv.CurrentRow.Index
            str1 = dgv("CHARTYPE", intRow).Value
            If StrComp(str1, "RC", CompareMethod.Text) = 0 Then
                dgv.Columns("NUMWATSONRUNNUMBER").Visible = True
            Else
                dgv.Columns("NUMWATSONRUNNUMBER").Visible = False
            End If

        Catch ex As Exception

        End Try

    End Sub

End Class