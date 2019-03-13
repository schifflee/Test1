Option Compare Text

Public Class frmSDHome

    Public boolHold As Boolean = False
    Public boolSourceSD As Boolean
    Public boolFormLoad As Boolean = False
    Public boolCmdEditE As Boolean
    Public intRowEditP As Integer = 0
    Public intRowEditS As Integer = 0
    Public boolSDProjAdd As Boolean = False
    Public boolSDStudyAdd As Boolean = False
    Public boolFromWeekRange As Boolean = False
    Public boolLL As Boolean = False
    Public boolUL As Boolean = False
    Public SDProjAddID As Int64 = 0
    Public SDStudyAddID As Int64 = 0
    Public arrC(10, 100)
    '1=ColumnName,2=HeaderText,3=datatype,4=value for id's, 5=id for id's
    Public tblProj as New System.Data.Datatable
    Public tblStud as New System.Data.Datatable
    Public tblAss as New System.Data.Datatable
    Public tblRoute as New System.Data.Datatable
    Public tblCal as New System.Data.Datatable
    Public tblGroupSummary as New System.Data.Datatable
    Public StudProjCellValCol As String = "NA"
    Public dtG As Date
    Public dgvA As DataGridView
    Public boolFromcbxStudy As Boolean = False
    Public boolFromGroupSummary As Boolean = False
    Public boolFromGroupRoute As Boolean = False
    Public boolFromRouteRemove As Boolean = False
    Public boolFromApplyGroup As Boolean = False
    Public boolAssayCancel As Boolean = False


    Public aaa As Short


    Sub LoadCalendardgv()

        'Dim dgv As DataGridView
        'Dim dv as system.data.dataview
        'Dim tbl as System.Data.Datatable

        'tbl = frmH.QRYGUWUCALENDAR
        'dv = New DataView(tbl)

        'dgv = Me.dgvCalendar
        'dgv.DataSource = dv


    End Sub

    Sub CreateGroupSummary()

        Dim dtbl as System.Data.Datatable
        Dim rows1() As DataRow
        Dim rows2() As DataRow
        Dim strF As String
        Dim strF1 As String
        Dim strF2 As String
        Dim strS As String
        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView
        Dim intRow As Short
        Dim id As Int64
        Dim intRows As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim var1, var2
        Dim dv1 as system.data.dataview
        Dim dv2 as system.data.dataview
        Dim dv3 as system.data.dataview
        Dim intRowsdv1 As Short
        Dim intRowsdv2 As Short
        Dim id1 As Int64
        Dim id2 As Int64
        Dim str1 As String
        Dim str2 As String

        dtbl = Me.tblGroupSummary
        dtbl.Clear()

        If dtbl.Columns.Count > 0 Then
        Else
            Dim col1 As New DataColumn
            col1.ColumnName = "ID_TBLGUWUPKGROUPS"
            dtbl.Columns.Add(col1)

            Dim col2 As New DataColumn
            col2.ColumnName = "ID_TBLGUWUPKROUTES"
            dtbl.Columns.Add(col2)

            Dim col3 As New DataColumn
            col3.ColumnName = "ColumnValue"
            dtbl.Columns.Add(col3)
        End If

        dgv1 = Me.dgvAssays

        If dgv1.Rows.Count = 0 Then
            intRow = -1
            id = -1
        ElseIf dgv1.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = dgv1.CurrentRow.Index
        End If
        If intRow = -1 Then
        Else
            id = dgv1("ID_TBLGUWUASSAY", intRow).Value
        End If

        strF = "ID_TBLGUWUASSAY = " & id

        dv1 = New DataView(tblGuWuPKGroups)
        dv1.RowFilter = strF
        intRowsdv1 = dv1.Count

        dv2 = New DataView(tblGuWuPKRoutes)

        For Count1 = 0 To intRowsdv1 - 1

            id1 = dv1.Item(Count1).Item("ID_TBLGUWUPKGROUPS")

            Dim nrow As DataRow = dtbl.NewRow
            nrow.BeginEdit()
            var1 = id1 ' -1 'rows1(Count1).Item("ID_TBLGUWUPKGROUPS")
            var2 = -1 'rows1(Count1).Item("ID_TBLGUWUPKROUTES")
            nrow("ID_TBLGUWUPKGROUPS") = var1
            nrow("ID_TBLGUWUPKROUTES") = var2
            nrow("ColumnValue") = dv1(Count1).Item("CHARGROUP")
            nrow.EndEdit()
            dtbl.Rows.Add(nrow)

            strF2 = "ID_TBLGUWUPKGROUPS = " & id1
            dv2.RowFilter = strF2
            dv2.Sort = "CHARROUTE ASC"
            intRowsdv2 = dv2.Count
            For Count2 = 0 To intRowsdv2 - 1
                Dim nrow1 As DataRow = dtbl.NewRow
                nrow1.BeginEdit()
                var1 = dv2(Count2).Item("ID_TBLGUWUPKGROUPS")
                var2 = dv2(Count2).Item("ID_TBLGUWUPKROUTES")
                nrow1("ID_TBLGUWUPKGROUPS") = var1
                nrow1("ID_TBLGUWUPKROUTES") = var2
                nrow1("ColumnValue") = "     " & dv2(Count2).Item("CHARROUTE")
                nrow1.EndEdit()
                dtbl.Rows.Add(nrow1)
            Next
        Next

        dv3 = New DataView(dtbl)

        dgv2 = Me.dgvGroupSummary
        dv3.AllowDelete = False
        dv3.AllowNew = False
        dv3.AllowEdit = False

        dgv2.DataSource = dv3

        dgv2.ReadOnly = True
        dgv2.RowHeadersWidth = 10

        For Count1 = 0 To dgv2.Columns.Count - 1
            str1 = dgv2.Columns(Count1).Name
            If StrComp(str1, "ColumnValue", CompareMethod.Text) = 0 Then
                dgv2.Columns(Count1).Visible = True
                dgv2.Columns(Count1).HeaderText = "Group/" & ChrW(10) & "Route"
                dgv2.Columns(Count1).SortMode = DataGridViewColumnSortMode.NotSortable

            Else
                dgv2.Columns(Count1).Visible = False
            End If
        Next

        dgv2.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
        dgv2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        dgv2.AllowUserToResizeColumns = True
        dgv2.AllowUserToResizeRows = True
        dgv2.RowHeadersWidth = 10
        dgv2.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        dgv2.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dgv2.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True


        dgv2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        Dim mw As Int16

        mw = dgv2.Width - (dgv2.RowHeadersWidth * 1.2)

        'dgv2.Columns.Item("ColumnValue").MinimumWidth = 150 'mw
        'dgv2.Columns.Item("ColumnValue").MinimumWidth = mw
        'dgv2.Columns("ColumnValue").MinimumWidth = mw
        Try
            dgv2.Columns.Item("ColumnValue").MinimumWidth = mw
        Catch ex As Exception

        End Try
        Try
            dgv2.Columns.Item("ColummValue").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        Catch ex As Exception

        End Try

        dgv2.AutoResizeColumns()

        Call UpdateGroupSummarySelection()

    End Sub

    Sub LockAll(ByVal bool As Boolean)
        Call LockSDProjects(bool)
        Call LockSDStudies(bool)
        Call LockCPTab(bool)
        Call LockAssays(bool)

    End Sub

    Sub LockSDProjects(ByVal bool)

        'Dim dv as system.data.dataview

        'dv = Me.dgvSDProject.DataSource
        'dv.AllowEdit = Not (bool)

        Me.dgvSDProject.ReadOnly = True
        Me.dgvProj.ReadOnly = bool

        Me.cmdAddProject.Enabled = Not (bool)

    End Sub

    Sub LockSDStudies(ByVal bool)

        'Dim dv as system.data.dataview

        'dv = Me.dgvSDStudy.DataSource
        'dv.AllowEdit = Not (bool)

        Me.dgvSDStudy.ReadOnly = True
        Me.dgvStud.ReadOnly = bool

        Me.cmdAddStudy.Enabled = Not (bool)
        Me.cmdWatsonExport.Enabled = Not (bool)


    End Sub

    Sub LockAssays(ByVal bool)

        Me.dgvAssays.ReadOnly = True
        Me.dgvAss.ReadOnly = bool
        Me.dgvGroups.ReadOnly = bool
        Me.dgvRoutes.ReadOnly = bool
        Me.dgvGroupDetails.ReadOnly = bool
        Me.lbxRoute.Enabled = Not (bool)

        Me.dgvGroupSummary.ReadOnly = True
        Me.dgvCmpd.ReadOnly = True
        Me.dgvLotNum.ReadOnly = True
        Me.dgvPI.ReadOnly = True
        Me.dgvAnalyst.ReadOnly = True
        Me.dgvGroupTimePoints.ReadOnly = True
        Me.dgvPatients.ReadOnly = True

        Me.cmdAddAssay.Enabled = Not (bool)
        Me.cmdAddGroup.Enabled = Not (bool)
        Me.cmdRemoveGroup.Enabled = Not (bool)
        Me.cmdAddRoute.Enabled = Not (bool)
        Me.cmdRemoveRoute.Enabled = Not (bool)
        Me.cmdApplyToAllGroups.Enabled = Not (bool)
        Me.cmdGetAssay.Enabled = Not (bool)
        Me.cmdAddCompound.Enabled = Not (bool)
        Me.cmdGetLotNum.Enabled = Not (bool)
        Me.cmdConfigPI.Enabled = Not (bool)
        Me.cmdConfigAnalyst.Enabled = Not (bool)

        Me.cmdTimePoints.Enabled = Not (bool)
        Me.cmdPatients.Enabled = Not (bool)


    End Sub

    Private Sub frmSDHome_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        Dim str1 As String
        Dim str2 As String

        boolFormLoad = True

        Call FilllbxSymbol(Me.lbxSymbol, Me.lblSymbol1)

        Call ShowGuWudgv()

        Me.cmdEdit.Enabled = True

        str2 = GetVersion()

        'record guest user
        str1 = "LABIntegrity StudyDoc" & ChrW(174) & " - Study Designer"
        str1 = str1 & " v" & str2 & gUserLabel ' " - User: Guest"

        Me.Text = str1

        Call LoadCalendardgv()

        'Me.SpinEdit1.Value = 5

        boolFormLoad = False

    End Sub

    Private Sub lbxSymbol_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbxSymbol.SelectedIndexChanged
        Dim var1

        var1 = Me.lbxSymbol.SelectedItem
        If StrComp(var1, "nbh", CompareMethod.Text) = 0 Then
            var1 = ChrW(2011) 'NBH
        ElseIf StrComp(var1, "nbsp", CompareMethod.Text) = 0 Then
            var1 = ChrW(160)
        ElseIf StrComp(var1, "CR", CompareMethod.Text) = 0 Then
            var1 = ChrW(10)
        Else
        End If
        Me.txtSymbol.Select()
        Me.txtSymbol.Text = var1

        Me.txtSymbol.SelectAll()

        'SendKeys.Send("+{END}")

    End Sub

    Sub ForceCellFormat(ByVal dgv As DataGridView)
        Dim str1 As String
        Dim str2 As String
        Dim intRow As Short
        'Dim dgv As DataGridView
        Dim strBool As String
        Dim strDt As String
        Dim locX, locY
        Dim var1, var2, var3
        Dim rows() As DataRow
        Dim strF As String
        Dim strS As String
        Dim boolRO As Boolean
        Dim int1 As Short
        Dim Count1 As Short

        dgvA = dgv

        'dgv = Me.dgv1
        Me.mCal1.Visible = False

        intRow = dgv.CurrentRow.Index
        str1 = dgv.Rows(intRow).Cells("ColumnName").Value
        boolRO = dgv.Rows(intRow).Cells("ColumnValue").ReadOnly
        If boolRO Then
            GoTo end1
        End If

        'aaa = aaa + 1
        'Me.lbx1.Items.Add(aaa & " - SelectionChanged - " & str1)
        'Me.lbx1.SelectedIndex = Me.lbx1.Items.Count - 1


        strBool = Mid(str1, 1, 4)
        strDt = Mid(str1, 1, 2)
        var1 = NZ(dgv.Rows(intRow).Cells("ColumnValue").Value, "")
        If StrComp(strBool, "BOOL", CompareMethod.Text) = 0 Then
            'make cell a checkbox
            'Dim chk1 As New DataGridViewCheckBoxCell
            'chk1.Value = 0
            'dgv1("ColumnValue", intRow) = chk1

            Dim cbx As New DataGridViewComboBoxCell
            cbx.AutoComplete = True
            cbx.MaxDropDownItems = 20
            cbx.Items.Add("TRUE")
            cbx.Items.Add("FALSE")
            If Len(var1) = 0 Then
                cbx.Value = "FALSE"
            Else
                cbx.Value = var1
            End If

            cbx.DisplayStyleForCurrentCellOnly = True
            cbx.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
            dgv("ColumnValue", intRow) = cbx

        ElseIf StrComp(strDt, "DT", CompareMethod.Text) = 0 Then

            Select Case dgv.Name
                Case "dgvAss"
                    locX = Me.sst1.Left + dgv.Left + dgv.RowHeadersWidth + (dgv.Columns(1).Width * 2)
                    locY = Me.sst1.Top + dgv.Location.Y + (dgv.Rows(intRow).Height * intRow) + dgv.ColumnHeadersHeight
                Case Else
                    locX = Me.sst1.Left + Me.tabStudies.Left + dgv.Left + dgv.RowHeadersWidth + (dgv.Columns(1).Width * 2)
                    locY = Me.sst1.Top + Me.tabStudies.Top + dgv.Location.Y + (dgv.Rows(intRow).Height * intRow) + dgv.ColumnHeadersHeight
            End Select

            If IsDate(var1) Then
                Me.mCal1.SelectionStart = var1
                Me.mCal1.SelectionEnd = var1
            Else
                Me.mCal1.SelectionStart = Now
                Me.mCal1.SelectionEnd = Now
            End If
            Me.mCal1.Location = new system.drawing.point(locX, locY)
            Me.mCal1.ScrollChange = 1
            Me.mCal1.MaxSelectionCount = 1

            Me.mCal1.Visible = True

        ElseIf StrComp(str1, "ID_TBLGUWUPROJECTS", CompareMethod.Text) = 0 Then

            dgv.Item("ColumnValue", intRow).ReadOnly = True

        ElseIf StrComp(str1, "ID_TBLCONFIGREPORTTYPE", CompareMethod.Text) = 0 Then
            Dim cbx As New DataGridViewComboBoxCell
            cbx.AutoComplete = True
            cbx.MaxDropDownItems = 20
            strF = "BOOLINCLUDE = -1"
            strS = "ID_TBLCONFIGREPORTTYPE ASC"
            rows = tblConfigReportType.Select(strF, strS)
            cbx.DataSource = rows
            cbx.DisplayMember = tblConfigReportType.Columns("CHARREPORTTYPE").ColumnName
            If Len(var1) = 0 Then
            Else
                cbx.Value = var1
            End If
            cbx.DisplayStyleForCurrentCellOnly = True
            cbx.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
            dgv("ColumnValue", intRow) = cbx

        ElseIf StrComp(str1, "ID_TBLGUWUSTUDYSTAT", CompareMethod.Text) = 0 Then
            Dim cbx As New DataGridViewComboBoxCell
            cbx.AutoComplete = True
            cbx.MaxDropDownItems = 20
            'strF = "BOOLINCLUDE = -1"
            'strS = "ID_TBLGUWUSTUDYSTAT ASC"
            'rows = tblGuWuStudyStat.Select(strF, strS)
            cbx.DataSource = tblGuWuStudyStat
            cbx.DisplayMember = tblGuWuStudyStat.Columns("CHARSTATUS").ColumnName
            If Len(var1) = 0 Then
            Else
                cbx.Value = var1
            End If
            cbx.DisplayStyleForCurrentCellOnly = True
            cbx.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
            dgv("ColumnValue", intRow) = cbx


        ElseIf StrComp(str1, "ID_TBLGUWUSTUDYDESIGNTYPE", CompareMethod.Text) = 0 Then
            Dim cbx As New DataGridViewComboBoxCell
            cbx.AutoComplete = True
            cbx.MaxDropDownItems = 20
            'strF = "BOOLINCLUDE = -1"
            'strS = "ID_TBLGUWUSTUDYDESIGNTYPE ASC"
            'rows = tblGuWuStudyDesignType.Select(strF, strS)
            cbx.DataSource = tblGuWuStudyDesignType
            cbx.DisplayMember = tblGuWuStudyDesignType.Columns("CHARSTUDYDESIGNTYPE").ColumnName
            If Len(var1) = 0 Then
            Else
                cbx.Value = var1
            End If
            cbx.DisplayStyleForCurrentCellOnly = True
            cbx.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
            dgv("ColumnValue", intRow) = cbx

        ElseIf StrComp(str1, "CHARSPECIES", CompareMethod.Text) = 0 Then

            Dim cbx As New DataGridViewComboBoxCell
            cbx.AutoComplete = True
            cbx.MaxDropDownItems = 20
            strF = "ID_TBLGUWUSPECIES > -1"
            'strS = "CHARGUWUSPECIES ASC"
            'rows = tblGuWuStudyDesignType.Select(strF, strS)
            'strS = "CHARGUWUSPECIES ASC"
            'rows = tblGuWuSpecies.Select(strF, strS)
            strS = "CHARSPECIES ASC"
            rows = tblGuWuSpecies.Select(strF, strS)
            cbx.DataSource = rows 'tblGuWuSpecies
            cbx.DisplayMember = tblGuWuSpecies.Columns("CHARSPECIES").ColumnName
            If Len(var1) = 0 Then
            Else
                cbx.Value = var1
            End If
            cbx.DisplayStyleForCurrentCellOnly = True
            cbx.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
            dgv("ColumnValue", intRow) = cbx

        ElseIf StrComp(str1, "CHARSPECIESSTRAIN", CompareMethod.Text) = 0 Then
            Dim cbx As New DataGridViewComboBoxCell
            cbx.AutoComplete = True
            cbx.MaxDropDownItems = 20
            int1 = FindRow(Me.dgvAss, "CHARSPECIES")
            var2 = Me.dgvAss("ColumnValue", int1).Value
            strF = "ID_TBLGUWUSPECIES = -1"
            If Len(NZ(var2, "")) = 0 Then
                strF = "ID_TBLGUWUSPECIES = -1"
            Else
                For Count1 = 0 To tblGuWuSpecies.Rows.Count - 1
                    var3 = tblGuWuSpecies.Rows(Count1).Item("CHARSPECIES")
                    If var2 = var3 Then
                        strF = "ID_TBLGUWUSPECIES = " & tblGuWuSpecies.Rows(Count1).Item("ID_TBLGUWUSPECIES")
                        Exit For
                    End If
                Next
            End If
            strS = "CHARSTRAIN ASC"
            rows = tblGuWuSpeciesStrain.Select(strF, strS)
            cbx.DataSource = rows 'tblGuWuSpeciesStrain
            cbx.DisplayMember = tblGuWuSpeciesStrain.Columns("CHARSTRAIN").ColumnName
            If Len(var1) = 0 Then
            Else
                cbx.Value = var1
            End If
            cbx.DisplayStyleForCurrentCellOnly = True
            cbx.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
            dgv("ColumnValue", intRow) = cbx

        ElseIf StrComp(str1, "CHARROUTE", CompareMethod.Text) = 0 Then

            dgv.Item("ColumnValue", intRow).ReadOnly = True

        ElseIf StrComp(str1, "CHARDOSEUNITS", CompareMethod.Text) = 0 Then
            Dim cbx As New DataGridViewComboBoxCell
            cbx.AutoComplete = True
            cbx.MaxDropDownItems = 20
            'strF = "BOOLINCLUDE = -1"
            'strS = "ID_TBLGUWUSTUDYDESIGNTYPE ASC"
            'rows = tblGuWuStudyDesignType.Select(strF, strS)
            cbx.DataSource = tblGuWuDoseUnits
            cbx.DisplayMember = tblGuWuDoseUnits.Columns("CHARDOSEUNITS").ColumnName
            If Len(var1) = 0 Then
            Else
                cbx.Value = var1
            End If
            cbx.DisplayStyleForCurrentCellOnly = True
            cbx.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
            dgv("ColumnValue", intRow) = cbx

        ElseIf StrComp(str1, "CHARDOSECONCUNITS", CompareMethod.Text) = 0 Then
            Dim cbx As New DataGridViewComboBoxCell
            cbx.AutoComplete = True
            cbx.MaxDropDownItems = 20
            'strF = "BOOLINCLUDE = -1"
            'strS = "ID_TBLGUWUSTUDYDESIGNTYPE ASC"
            'rows = tblGuWuStudyDesignType.Select(strF, strS)
            cbx.DataSource = tblGuWuDoseUnits
            cbx.DisplayMember = tblGuWuDoseUnits.Columns("CHARDOSEUNITS").ColumnName
            If Len(var1) = 0 Then
            Else
                cbx.Value = var1
            End If
            cbx.DisplayStyleForCurrentCellOnly = True
            cbx.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
            dgv("ColumnValue", intRow) = cbx

        ElseIf StrComp(str1, "CHARSTUDYDESIGNTYPE", CompareMethod.Text) = 0 Then
            Dim cbx As New DataGridViewComboBoxCell
            cbx.AutoComplete = True
            cbx.MaxDropDownItems = 20
            'strF = "BOOLINCLUDE = -1"
            'strS = "ID_TBLGUWUSTUDYDESIGNTYPE ASC"
            'rows = tblGuWuStudyDesignType.Select(strF, strS)
            cbx.DataSource = tblGuWuStudyDesignType
            cbx.DisplayMember = tblGuWuStudyDesignType.Columns("CHARSTUDYDESIGNTYPE").ColumnName
            If Len(var1) = 0 Then
            Else
                cbx.Value = var1
            End If
            cbx.DisplayStyleForCurrentCellOnly = True
            cbx.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
            dgv("ColumnValue", intRow) = cbx

        ElseIf StrComp(str1, "CHARVEHICLE", CompareMethod.Text) = 0 Then
            Dim cbx As New DataGridViewComboBoxCell
            cbx.AutoComplete = True
            cbx.MaxDropDownItems = 20
            strF = "ID_TBLDROPDOWNBOXNAME = 13"
            strS = "INTORDER ASC"
            rows = tblDropdownBoxContent.Select(strF, strS)
            cbx.DataSource = rows
            cbx.DisplayMember = tblDropdownBoxContent.Columns("CHARVALUE").ColumnName
            If Len(var1) = 0 Then
            Else
                cbx.Value = var1
            End If
            cbx.DisplayStyleForCurrentCellOnly = True
            cbx.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
            dgv("ColumnValue", intRow) = cbx

        ElseIf StrComp(str1, "CHARFORMULATION", CompareMethod.Text) = 0 Then
            Dim cbx As New DataGridViewComboBoxCell
            cbx.AutoComplete = True
            cbx.MaxDropDownItems = 20
            strF = "ID_TBLDROPDOWNBOXNAME = 14"
            strS = "INTORDER ASC"
            rows = tblDropdownBoxContent.Select(strF, strS)
            cbx.DataSource = rows
            cbx.DisplayMember = tblDropdownBoxContent.Columns("CHARVALUE").ColumnName
            If Len(var1) = 0 Then
            Else
                cbx.Value = var1
            End If
            cbx.DisplayStyleForCurrentCellOnly = True
            cbx.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
            dgv("ColumnValue", intRow) = cbx

        ElseIf StrComp(str1, "CHARREGIMEN", CompareMethod.Text) = 0 Then
            Dim cbx As New DataGridViewComboBoxCell
            cbx.AutoComplete = True
            cbx.MaxDropDownItems = 20
            strF = "ID_TBLDROPDOWNBOXNAME = 15"
            strS = "INTORDER ASC"
            rows = tblDropdownBoxContent.Select(strF, strS)
            cbx.DataSource = rows
            cbx.DisplayMember = tblDropdownBoxContent.Columns("CHARVALUE").ColumnName
            If Len(var1) = 0 Then
            Else
                cbx.Value = var1
            End If
            cbx.DisplayStyleForCurrentCellOnly = True
            cbx.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
            dgv("ColumnValue", intRow) = cbx

        ElseIf StrComp(str1, "CHARFASTED", CompareMethod.Text) = 0 Then
            Dim cbx As New DataGridViewComboBoxCell
            cbx.AutoComplete = True
            cbx.MaxDropDownItems = 20
            strF = "ID_TBLDROPDOWNBOXNAME = 16"
            strS = "INTORDER ASC"
            rows = tblDropdownBoxContent.Select(strF, strS)
            cbx.DataSource = rows
            cbx.DisplayMember = tblDropdownBoxContent.Columns("CHARVALUE").ColumnName
            If Len(var1) = 0 Then
            Else
                cbx.Value = var1
            End If
            cbx.DisplayStyleForCurrentCellOnly = True
            cbx.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
            dgv("ColumnValue", intRow) = cbx

        ElseIf StrComp(str1, "CHARRESTRAINED", CompareMethod.Text) = 0 Then
            Dim cbx As New DataGridViewComboBoxCell
            cbx.AutoComplete = True
            cbx.MaxDropDownItems = 20
            strF = "ID_TBLDROPDOWNBOXNAME = 17"
            strS = "INTORDER ASC"
            rows = tblDropdownBoxContent.Select(strF, strS)
            cbx.DataSource = rows
            cbx.DisplayMember = tblDropdownBoxContent.Columns("CHARVALUE").ColumnName
            If Len(var1) = 0 Then
            Else
                cbx.Value = var1
            End If
            cbx.DisplayStyleForCurrentCellOnly = True
            cbx.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
            dgv("ColumnValue", intRow) = cbx

        ElseIf StrComp(str1, "CHARGENDER", CompareMethod.Text) = 0 Then
            Dim cbx As New DataGridViewComboBoxCell
            cbx.AutoComplete = True
            cbx.MaxDropDownItems = 20
            strF = "ID_TBLDROPDOWNBOXNAME = 18"
            strS = "INTORDER ASC"
            rows = tblDropdownBoxContent.Select(strF, strS)
            cbx.DataSource = rows
            cbx.DisplayMember = tblDropdownBoxContent.Columns("CHARVALUE").ColumnName
            If Len(var1) = 0 Then
            Else
                cbx.Value = var1
            End If
            cbx.DisplayStyleForCurrentCellOnly = True
            cbx.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
            dgv("ColumnValue", intRow) = cbx

        ElseIf StrComp(str1, "CHARMATRIX", CompareMethod.Text) = 0 Then
            Dim cbx As New DataGridViewComboBoxCell
            cbx.AutoComplete = True
            cbx.MaxDropDownItems = 20
            strF = "ID_TBLDROPDOWNBOXNAME = 19"
            strS = "INTORDER ASC"
            rows = tblDropdownBoxContent.Select(strF, strS)
            cbx.DataSource = rows
            cbx.DisplayMember = tblDropdownBoxContent.Columns("CHARVALUE").ColumnName
            If Len(var1) = 0 Then
            Else
                cbx.Value = var1
            End If
            cbx.DisplayStyleForCurrentCellOnly = True
            cbx.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
            dgv("ColumnValue", intRow) = cbx

        ElseIf StrComp(str1, "CHARTISSUE", CompareMethod.Text) = 0 Then
            Dim cbx As New DataGridViewComboBoxCell
            cbx.AutoComplete = True
            cbx.MaxDropDownItems = 20
            strF = "ID_TBLDROPDOWNBOXNAME = 20"
            strS = "INTORDER ASC"
            rows = tblDropdownBoxContent.Select(strF, strS)
            cbx.DataSource = rows
            cbx.DisplayMember = tblDropdownBoxContent.Columns("CHARVALUE").ColumnName
            If Len(var1) = 0 Then
            Else
                cbx.Value = var1
            End If
            cbx.DisplayStyleForCurrentCellOnly = True
            cbx.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
            dgv("ColumnValue", intRow) = cbx

        ElseIf StrComp(str1, "CHARDOSUNITS", CompareMethod.Text) = 0 Then
            Dim cbx As New DataGridViewComboBoxCell
            cbx.AutoComplete = True
            cbx.MaxDropDownItems = 20
            strF = "ID_TBLDROPDOWNBOXNAME = 21"
            strS = "INTORDER ASC"
            rows = tblDropdownBoxContent.Select(strF, strS)
            cbx.DataSource = rows
            cbx.DisplayMember = tblDropdownBoxContent.Columns("CHARVALUE").ColumnName
            If Len(var1) = 0 Then
            Else
                cbx.Value = var1
            End If
            cbx.DisplayStyleForCurrentCellOnly = True
            cbx.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
            dgv("ColumnValue", intRow) = cbx

        ElseIf StrComp(str1, "CHARDOSECONCUNITS", CompareMethod.Text) = 0 Then
            Dim cbx As New DataGridViewComboBoxCell
            cbx.AutoComplete = True
            cbx.MaxDropDownItems = 20
            strF = "ID_TBLDROPDOWNBOXNAME = 22"
            strS = "INTORDER ASC"
            rows = tblDropdownBoxContent.Select(strF, strS)
            cbx.DataSource = rows
            cbx.DisplayMember = tblDropdownBoxContent.Columns("CHARVALUE").ColumnName
            If Len(var1) = 0 Then
            Else
                cbx.Value = var1
            End If
            cbx.DisplayStyleForCurrentCellOnly = True
            cbx.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
            dgv("ColumnValue", intRow) = cbx

        ElseIf StrComp(str1, "CHARTISSUEWTUNITS", CompareMethod.Text) = 0 Then
            Dim cbx As New DataGridViewComboBoxCell
            cbx.AutoComplete = True
            cbx.MaxDropDownItems = 20
            strF = "ID_TBLDROPDOWNBOXNAME = 23"
            strS = "INTORDER ASC"
            rows = tblDropdownBoxContent.Select(strF, strS)
            cbx.DataSource = rows
            cbx.DisplayMember = tblDropdownBoxContent.Columns("CHARVALUE").ColumnName
            If Len(var1) = 0 Then
            Else
                cbx.Value = var1
            End If
            cbx.DisplayStyleForCurrentCellOnly = True
            cbx.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
            dgv("ColumnValue", intRow) = cbx


        End If

end1:

    End Sub

    Private Sub cmdExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        'Dim var1
        'var1 = MsgBox("Do you wish to exit?", MsgBoxStyle.YesNo, "Exit GuWu Study Designer...")
        'Me.Refresh()
        'If var1 = 6 Then 'Yes then

        Call DoThis("cmdExit")

        Me.Visible = False

        frmC.Visible = True

        'ElseIf var1 = 7 Then

        'End If

    End Sub

    Sub FormLoad()

        boolFormLoad = True

        Me.rbMonthS.Checked = True

        Call FilllbxRoutes()

        'Call CreateCalendar()

        Call FillTab1()

        frmSD = Me

        Me.rbGuWu.Checked = True
        boolSourceSD = True

        Call ConfigureDGVs()
        Call ShowGuWudgv()

        Call Configdgv1(Me.dgvProj, Me.dgvSDProject, "Projects")
        Call Configdgv1(Me.dgvStud, Me.dgvSDStudy, "Studies")
        Call Configdgv1(Me.dgvAss, Me.dgvAssays, "Assays")
        Call Configdgv1(Me.dgvGroupDetails, Me.dgvRoutes, "GroupDetails")

        Call CPTab_Initialize()

        Call SDDataSourceChecked()

        Call FillSchedFilter()

        Call LoadScbx()

        Call DoAllLabels()

        Me.cbxSchedFilter.SelectedIndex = 0
        'Me.SchedulerControl1.MonthView.WeekCount = 5

        Call LockAll(True)

        'pesky
        Try
            Me.dgvGroupTimePoints.Columns("NUMTIMEPOINT").SortMode = DataGridViewColumnSortMode.NotSortable
        Catch ex As Exception
            Dim str1 As String
            str1 = ""
        End Try
        Try
            Me.dgvPatients.Columns("CHARSUBJECTNAME").SortMode = DataGridViewColumnSortMode.NotSortable
            Me.dgvPatients.Columns("BOOLSERIALBLEED").SortMode = DataGridViewColumnSortMode.NotSortable
            Me.dgvPatients.Columns("BOOLTERMINALBLEED").SortMode = DataGridViewColumnSortMode.NotSortable
            Me.dgvPatients.Columns("NUMTIMEPOINT").SortMode = DataGridViewColumnSortMode.NotSortable
        Catch ex As Exception
            Dim str1 As String
            str1 = ""
        End Try

        boolCmdEditE = True


        boolFormLoad = False

    End Sub

    Sub FillTab1()

        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim strF As String
        Dim strS As String
        Dim Count1 As Short
        Dim str1 As String


        strF = "CHARFORM = 'Study Design' AND INTFORM = 6 AND BOOLINCLUDEINTEMPLATE = -1"
        strS = "INTORDER ASC"
        dv = New DataView(tblTab1, strF, strS, DataViewRowState.CurrentRows)
        dv.AllowDelete = False
        dv.AllowEdit = False
        dv.AllowNew = False

        dgv = Me.dgvTab1
        dgv.DataSource = dv

        'show only CHARFORM
        For Count1 = 0 To dgv.ColumnCount - 1
            dgv.Columns(Count1).Visible = False
        Next

        dgv.Columns("CHARITEM").Visible = True
        dgv.Columns("CHARITEM").HeaderText = ""
        dgv.RowHeadersWidth = 25

        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

    End Sub

    Sub SaveCP()

        Dim dv As system.data.dataview
        Dim tbl As System.Data.Datatable
        Dim dtbl As System.Data.Datatable
        Dim tbl1 As System.Data.Datatable
        Dim tbl2 As System.Data.Datatable
        Dim strF As String
        Dim ct1 As Short
        Dim ct2 As Short
        Dim ct3 As Short
        Dim drows1() As DataRow
        Dim dv2 As system.data.dataview
        Dim row As DataRow
        Dim dr() As DataRow
        Dim boolExists As Boolean
        Dim Count1 As Short
        Dim Count2 As Short
        Dim var1, var2, var3
        Dim int1 As Short
        Dim col As DataColumn
        Dim str1 As String
        Dim dv1 As system.data.dataview

        Me.dgvContributingPersonnel.CommitEdit(DataGridViewDataErrorContexts.Commit)

        tbl1 = tblContributingPersonnel
        ct1 = tbl1.Rows.Count
        dv = Me.dgvContributingPersonnel.DataSource
        ct2 = dv.Count

        'determine if criteria is met
        'must have a name
        dv.AllowDelete = True
        For Count1 = ct2 - 1 To 0 Step -1
            If dv(Count1).Row.RowState = DataRowState.Deleted Then 'ignore
            Else
                var1 = NZ(dv(Count1).Item("charCPName"), "")
                If Len(var1) = 0 Then 'delete
                    dv(Count1).Row.Delete()
                End If
            End If
        Next
        dv.AllowDelete = False

        ''now fix intOrder values
        'str1 = "intOrder ASC"
        'dv.Sort = str1
        'ct2 = dv.Count

        'For Count1 = 0 To ct2 - 2
        '    var1 = NZ(dv(Count1).Item("intorder"), 0)
        '    If Count1 = 0 And var1 = 0 Then
        '        var1 = 1
        '        dv(Count1).Row.BeginEdit()
        '        dv(Count1).Item("intOrder") = var1
        '        dv(Count1).Row.EndEdit()
        '    End If
        '    var2 = NZ(dv(Count1 + 1).Item("intorder"), 0)
        '    If var1 + 1 = var2 Then 'proceed
        '    Else 'fix
        '        var2 = var1 + 1
        '        dv(Count1).Row.BeginEdit()
        '        dv(Count1 + 1).Item("intOrder") = var2
        '        dv(Count1 + 1).Row.EndEdit()
        '    End If
        'Next

        'for some reason, BOOLINCLUDESIGONTABLEPAGE is null instead of zero
        'For Count1 = 0 To ct2 - 1
        '    var1 = dv(Count1).Item("BOOLINCLUDESIGONTABLEPAGE")
        '    If IsDBNull(var1) Then
        '        dv(Count1).BeginEdit()
        '        dv(Count1).Item("BOOLINCLUDESIGONTABLEPAGE") = 0
        '        dv(Count1).EndEdit()
        '    End If
        'Next

        Dim dvCheck As system.data.dataview = New DataView(tblContributingPersonnel)
        dvCheck.RowStateFilter = DataViewRowState.ModifiedCurrent
        Dim int10 As Short
        int10 = 1
        If int10 = 0 Then
        Else
            'endif
            If boolGuWuOracle Then
                Try
                    ta_tblContributingPersonnel.Update(tblContributingPersonnel)
                Catch ex As DBConcurrencyException
                    'ds2005.TBLCONTRIBUTINGPERSONNEL.Merge('ds2005.TBLCONTRIBUTINGPERSONNEL, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_tblContributingPersonnelAcc.Update(tblContributingPersonnel)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLCONTRIBUTINGPERSONNEL.Merge('ds2005Acc.TBLCONTRIBUTINGPERSONNEL, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblContributingPersonnelSQLServer.Update(tblContributingPersonnel)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLCONTRIBUTINGPERSONNEL.Merge('ds2005Acc.TBLCONTRIBUTINGPERSONNEL, True)
                End Try
            End If

        End If


    End Sub

    Sub ShowGuWudgv()

        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView
        Dim dgv3 As DataGridView
        Dim boolGuWu As Boolean
        Dim w
        Dim boolTab As Boolean

        If Me.rbGuWu.Checked Then
            boolGuWu = True
        Else
            boolGuWu = False
        End If

        w = Me.sst1.Width

        'first do Projects
        dgv1 = Me.dgvwProject
        dgv2 = Me.dgvSDProject
        dgv3 = Me.dgvProj

        If Me.rbProjectTab.Checked Then
            boolTab = True
        Else
            boolTab = False
        End If

        dgv3.Width = w - dgv3.Left - 30
        If boolTab Then
            dgv2.Width = w - dgv2.Left - 30
            dgv3.Visible = False
        Else
            dgv2.Width = 140
            dgv3.Visible = True
        End If

        dgv1.Width = dgv2.Width
        If boolGuWu Then
            dgv1.Width = 140
        Else
            dgv1.Width = w - dgv1.Left - 30
            dgv3.Visible = False
        End If


        If boolGuWu Then
            dgv1.Visible = Not (boolGuWu)
            dgv2.Visible = boolGuWu
        Else
            dgv2.Visible = boolGuWu
            dgv1.Visible = Not (boolGuWu)
        End If

        dgv2.Left = dgv1.Left
        dgv2.Top = dgv1.Top

        dgv2.Height = dgv1.Height
        dgv2.Height = dgv1.Height

        dgv1 = Me.dgvwProjectS
        dgv2 = Me.dgvSDProjectS

        If boolGuWu Then
            dgv1.Visible = Not (boolGuWu)
            dgv2.Visible = boolGuWu
        Else
            dgv2.Visible = boolGuWu
            dgv1.Visible = Not (boolGuWu)
        End If

        dgv2.Left = dgv1.Left
        dgv2.Top = dgv1.Top

        dgv2.Height = dgv1.Height
        dgv2.Height = dgv1.Height


        'now do Studies

        If Me.rbStudyTab.Checked Then
            boolTab = True
        Else
            boolTab = False
        End If

        dgv1 = Me.dgvwStudy
        dgv2 = Me.dgvSDStudy
        dgv3 = Me.dgvSDProjectS


        If boolTab Then
            dgv2.Width = w - dgv2.Left - 30
            Me.tabStudies.Visible = False
        Else
            dgv2.Width = dgv3.Width
            Me.tabStudies.Visible = True
        End If

        If boolGuWu Then
            dgv1.Width = w - dgv1.Left - 30 '140
        Else
            dgv1.Width = w - dgv1.Left - 30
        End If

        If boolGuWu Then
            dgv1.Visible = Not (boolGuWu)
            dgv2.Visible = boolGuWu
        Else
            dgv2.Visible = boolGuWu
            dgv1.Visible = Not (boolGuWu)
            Me.tabStudies.Visible = False
        End If

        dgv2.Left = dgv1.Left
        dgv2.Top = dgv1.Top
        dgv2.Height = dgv1.Height
        dgv2.Height = dgv1.Height

        dgv2.Height = dgv1.Height
        dgv2.Height = dgv1.Height


        'If boolGuWu Then
        '    Me.dgvProj.Visible = True
        '    Me.tabStudies.Visible = True
        'Else
        '    Me.dgvProj.Visible = False
        '    Me.tabStudies.Visible = False
        'End If


    End Sub

    Sub Filldgv1GroupDetails()

        ''EG: dgv1=dgvProj, dvg2=dgvSDProj
        'Dim intRows As Short
        'Dim Count1 As Short
        'Dim str1 As String
        'Dim intRow As Short
        'Dim var1, var2
        'Dim strBool As String
        'Dim strF As String
        'Dim strS As String
        'Dim rows() As DataRow
        'Dim strDT As String
        'Dim strDt1 As String
        'Dim strDt2 As String
        'Dim dgv1 As DataGridView
        'Dim dgv2 As DataGridView
        'Dim dgv3 As DataGridView

        'dgv1 = Me.dgvGroupDetails
        'dgv2 = Me.dgvRoutes
        'dgv3 = Me.dgvGroups

        'intRows = dgv1.Rows.Count


        'var1 = dgv1.Name 'debugging
        'var2 = dgv2.Name 'debugging

        'If dgv2.CurrentRow Is Nothing Then
        '    For Count1 = 0 To intRows - 1
        '        dgv1("ColumnValue", Count1).Value = ""
        '        dgv1("ColumnValueActual", Count1).Value = ""
        '    Next
        '    Select Case strMod
        '        Case "Projects"

        '        Case "Studies"
        '            id_tblGuWuStudies = -1
        '    End Select

        'Else
        '    intRow = dgv2.CurrentRow.Index

        '    For Count1 = 0 To intRows - 1
        '        str1 = dgv1("ColumnName", Count1).Value
        '        strBool = Mid(str1, 1, 4)
        '        var1 = dgv2(str1, intRow).Value
        '        dgv1("ColumnValueActual", Count1).Value = var1
        '        If StrComp(strBool, "BOOL", CompareMethod.Text) = 0 Then
        '            If var1 = -1 Then
        '                var2 = "TRUE"
        '            Else
        '                var2 = "FALSE"
        '            End If
        '        ElseIf StrComp(str1, "ID_TBLGUWUPROJECTS", CompareMethod.Text) = 0 Then

        '            If Len(NZ(var1, "")) = 0 Then
        '                var2 = DBNull.Value
        '            Else
        '                dgv1.Item("ColumnValue", intRow).ReadOnly = True
        '                strF = "ID_TBLGUWUPROJECTS = " & var1
        '                rows = tblGuWuProjects.Select(strF)
        '                var2 = rows(0).Item("CHARPROJECTNUM")

        '            End If
        '            dgv1("ColumnReadOnly", Count1).Value = "TRUE"

        '        ElseIf StrComp(str1, "ID_TBLCONFIGREPORTTYPE", CompareMethod.Text) = 0 Then

        '            If Len(NZ(var1, "")) = 0 Then
        '                var2 = DBNull.Value
        '            Else
        '                strF = "ID_TBLCONFIGREPORTTYPE = " & var1
        '                rows = tblConfigReportType.Select(strF)
        '                var2 = rows(0).Item("CHARREPORTTYPE")
        '            End If

        '        ElseIf StrComp(str1, "ID_TBLGUWUSTUDYSTAT", CompareMethod.Text) = 0 Then

        '            If Len(NZ(var1, "")) = 0 Then
        '                var2 = DBNull.Value
        '            Else
        '                strF = "ID_TBLGUWUSTUDYSTAT = " & var1
        '                rows = tblGuWuStudyStat.Select(strF)
        '                var2 = rows(0).Item("CHARSTATUS")
        '            End If

        '        ElseIf StrComp(str1, "ID_TBLGUWUSTUDYDESIGNTYPE", CompareMethod.Text) = 0 Then

        '            If Len(NZ(var1, "")) = 0 Then
        '                var2 = DBNull.Value
        '            Else
        '                strF = "ID_TBLGUWUSTUDYDESIGNTYPE = " & var1
        '                rows = tblGuWuStudyDesignType.Select(strF)
        '                var2 = rows(0).Item("CHARSTUDYDESIGNTYPE")
        '            End If

        '        ElseIf IsUserIDStuff(str1) Then
        '            strDT = dgv1("ColumnDataType", Count1).Value
        '            'If InStr(strDT, "date", CompareMethod.Text) > 0 Then
        '            '    strDt1 = Format(var1, "long date")
        '            '    strDt2 = Format(var1, "long time")

        '            '    var2 = strDt1 & " " & strDt2
        '            '    dgv1("ColumnValueActual", Count1).Value = var2
        '            'Else
        '            '    var2 = var1
        '            'End If

        '            If Len(NZ(var1, "")) = 0 Then
        '                var2 = var1
        '            Else
        '                If InStr(strDT, "date", CompareMethod.Text) > 0 Then
        '                    strDt1 = Format(var1, "long date")
        '                    strDt2 = Format(var1, "long time")

        '                    var2 = strDt1 & " " & strDt2
        '                    dgv1("ColumnValueActual", Count1).Value = var2
        '                Else
        '                    var2 = var1
        '                End If
        '            End If
        '        Else
        '            var2 = var1
        '        End If

        '        dgv1("ColumnValue", Count1).Value = var2
        '    Next

        '    Select Case strMod
        '        Case "Projects"

        '        Case "Studies"
        '            id_tblGuWuStudies = dgv2.Item("ID_TBLGUWUSTUDIES", intRow).Value
        '    End Select

        'End If
    End Sub

    Sub Filldgv1(ByVal dgv1 As DataGridView, ByVal dgv2 As DataGridView, ByVal strMod As String)

        'EG: dgv1=dgvProj, dvg2=dgvSDProj
        Dim intRows As Short
        Dim Count1 As Short
        Dim str1 As String
        Dim intRow As Short
        Dim var1, var2
        Dim strBool As String
        Dim strF As String
        Dim strS As String
        Dim rows() As DataRow
        Dim strDT As String
        Dim strDt1 As String
        Dim strDt2 As String
        Dim varA, varB

        intRows = dgv1.Rows.Count
        Select Case strMod
            Case "Projects"

            Case "Studies"

        End Select

        varA = dgv1.Name 'debugging
        varB = dgv2.Name 'debugging

        If dgv1.Columns.Count = 0 Then
            Exit Sub
        End If

        If dgv2.CurrentRow Is Nothing Then
            For Count1 = 0 To intRows - 1
                dgv1("ColumnValue", Count1).Value = ""
                dgv1("ColumnValueActual", Count1).Value = ""
            Next
            Select Case strMod
                Case "Projects"

                Case "Studies"
                    id_tblGuWuStudies = -1
            End Select

        Else
            intRow = dgv2.CurrentRow.Index

            For Count1 = 0 To intRows - 1
                str1 = dgv1("ColumnName", Count1).Value
                strBool = Mid(str1, 1, 4)
                var1 = dgv2(str1, intRow).Value
                dgv1("ColumnValueActual", Count1).Value = var1
                If StrComp(strBool, "BOOL", CompareMethod.Text) = 0 Then
                    If var1 = -1 Then
                        var2 = "TRUE"
                    Else
                        var2 = "FALSE"
                    End If
                ElseIf StrComp(str1, "ID_TBLGUWUPROJECTS", CompareMethod.Text) = 0 Then

                    If Len(NZ(var1, "")) = 0 Then
                        var2 = DBNull.Value
                    Else
                        dgv1.Item("ColumnValue", intRow).ReadOnly = True
                        strF = "ID_TBLGUWUPROJECTS = " & var1
                        rows = tblGuWuProjects.Select(strF)
                        var2 = rows(0).Item("CHARPROJECTNUM")

                    End If
                    dgv1("ColumnReadOnly", Count1).Value = "TRUE"

                ElseIf StrComp(str1, "ID_TBLCONFIGREPORTTYPE", CompareMethod.Text) = 0 Then

                    If Len(NZ(var1, "")) = 0 Then
                        var2 = DBNull.Value
                    Else
                        strF = "ID_TBLCONFIGREPORTTYPE = " & var1
                        rows = tblConfigReportType.Select(strF)
                        var2 = NZ(rows(0).Item("CHARREPORTTYPE"), "Sample Analysis")
                    End If

                ElseIf StrComp(str1, "ID_TBLGUWUSTUDYSTAT", CompareMethod.Text) = 0 Then

                    If Len(NZ(var1, "")) = 0 Then
                        var2 = DBNull.Value
                    Else
                        strF = "ID_TBLGUWUSTUDYSTAT = " & var1
                        rows = tblGuWuStudyStat.Select(strF)
                        var2 = rows(0).Item("CHARSTATUS")
                    End If

                ElseIf StrComp(str1, "ID_TBLGUWUSTUDYDESIGNTYPE", CompareMethod.Text) = 0 Then

                    If Len(NZ(var1, "")) = 0 Then
                        var2 = DBNull.Value
                    Else
                        strF = "ID_TBLGUWUSTUDYDESIGNTYPE = " & var1
                        rows = tblGuWuStudyDesignType.Select(strF)
                        var2 = rows(0).Item("CHARSTUDYDESIGNTYPE")
                    End If

                ElseIf IsUserIDStuff(str1) Then
                    strDT = dgv1("ColumnDataType", Count1).Value
                    'If InStr(strDT, "date", CompareMethod.Text) > 0 Then
                    '    strDt1 = Format(var1, "long date")
                    '    strDt2 = Format(var1, "long time")

                    '    var2 = strDt1 & " " & strDt2
                    '    dgv1("ColumnValueActual", Count1).Value = var2
                    'Else
                    '    var2 = var1
                    'End If

                    If Len(NZ(var1, "")) = 0 Then
                        var2 = var1
                    Else
                        If InStr(strDT, "date", CompareMethod.Text) > 0 Then
                            strDt1 = Format(var1, "long date")
                            strDt2 = Format(var1, "long time")

                            var2 = strDt1 & " " & strDt2
                            dgv1("ColumnValueActual", Count1).Value = var2
                        Else
                            var2 = var1
                        End If
                    End If
                ElseIf StrComp(str1, "BOOLINCLUDE", CompareMethod.Text) = 0 Then
                    'change this later
                    var2 = -1
                ElseIf StrComp(str1, "BOOLACCEPTED", CompareMethod.Text) = 0 Then
                    'change this later
                    var2 = -1

                Else
                    var2 = var1
                End If

                dgv1("ColumnValue", Count1).Value = var2
            Next

            Select Case strMod
                Case "Projects"

                Case "Studies"
                    id_tblGuWuStudies = dgv2.Item("ID_TBLGUWUSTUDIES", intRow).Value
            End Select

        End If

    End Sub

    Sub Configdgv1(ByVal dgv1 As DataGridView, ByVal dgv2 As DataGridView, ByVal strMod As String)

        'e.g. Me.dgvProj, Me.dgvSDProject

        Dim intRows As Short
        Dim intCols As Short
        Dim int1 As Short
        Dim int2 As Short
        Dim var1, var2
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim dtbl As New System.Data.Datatable
        Dim bool As Boolean
        Dim strID As String
        Dim intRow As Short
        Dim tblT As System.Data.Datatable
        Dim mw As Int16

        'arrC
        '1=ColumnName,2=HeaderText,3=datatype,4=value for id's, 5=id for id's, 6=boolReadOnly

        'dgv1 = Me.dgv1
        intCols = dgv2.Columns.Count

        dgv2.ReadOnly = False

        dgv1.RowHeadersWidth = 25

        int1 = 0
        Select Case strMod
            Case "Projects"
                tblT = tblGuWuProjects
                dtbl = tblProj
                For Count1 = 0 To intCols - 1
                    If dgv2.Columns(Count1).Visible Then
                        bool = dgv2.Columns(Count1).ReadOnly
                        If bool Then
                        Else
                            int1 = int1 + 1
                            arrC(1, int1) = dgv2.Columns(Count1).Name
                            arrC(2, int1) = dgv2.Columns(Count1).HeaderText

                        End If
                    End If
                Next
                mw = 150

            Case "Studies"
                tblT = tblGuWuStudies
                dtbl = tblStud
                intRow = frmSD.dgvSDProjectS.CurrentRow.Index
                For Count1 = 0 To intCols - 1
                    str1 = dgv2.Columns(Count1).Name
                    If StrComp(str1, "ID_TBLGUWUSTUDIES", CompareMethod.Text) = 0 Then
                    Else
                        strID = Mid(str1, 1, 3)
                        If dgv2.Columns(Count1).Visible Or StrComp(strID, "ID_", CompareMethod.Text) = 0 Then
                            bool = dgv2.Columns(Count1).ReadOnly
                            If bool And StrComp(strID, "ID_TBLGUWUSTUDIES", CompareMethod.Text) <> 0 Then
                            Else
                                int1 = int1 + 1
                                arrC(1, int1) = dgv2.Columns(Count1).Name
                                arrC(2, int1) = dgv2.Columns(Count1).HeaderText
                                If StrComp(str1, "ID_TBLGUWUPROJECTS", CompareMethod.Text) = 0 Then
                                    'enter Project Number and ID
                                    var1 = frmSD.dgvSDProjectS.Item("CHARPROJECTNUM", intRow).Value
                                    var2 = frmSD.dgvSDProjectS.Item("ID_TBLGUWUPROJECTS", intRow).Value
                                    arrC(4, int1) = var1
                                    arrC(5, int1) = var2
                                End If
                            End If
                        End If
                    End If
                Next

                mw = 150

            Case "Assays"
                tblT = tblGuWuAssay
                dtbl = tblAss

                If frmSD.dgvSDStudy.Rows.Count = 0 Then
                    GoTo end1
                End If

                If frmSD.dgvSDStudy.CurrentRow Is Nothing Then
                    intRow = 0
                Else
                    intRow = frmSD.dgvSDStudy.CurrentRow.Index
                End If

                For Count1 = 0 To intCols - 1
                    str1 = dgv2.Columns(Count1).Name
                    If StrComp(str1, "ID_TBLGUWUASSAY", CompareMethod.Text) = 0 Or StrComp(str1, "ID_TBLGUWUSPECIES", CompareMethod.Text) = 0 Or StrComp(str1, "ID_TBLSTUDIES", CompareMethod.Text) = 0 Or StrComp(str1, "ID_TBLGUWUSTUDIES", CompareMethod.Text) = 0 Then
                    Else
                        strID = Mid(str1, 1, 3)
                        If dgv2.Columns(Count1).Visible Or StrComp(strID, "ID_", CompareMethod.Text) = 0 Then
                            bool = dgv2.Columns(Count1).ReadOnly
                            If bool And StrComp(strID, "ID_TBLGUWUASSAY", CompareMethod.Text) <> 0 Then
                            Else
                                int1 = int1 + 1
                                arrC(1, int1) = dgv2.Columns(Count1).Name
                                arrC(2, int1) = dgv2.Columns(Count1).HeaderText
                                'If StrComp(str1, "ID_TBLGUWUSTUDIES", CompareMethod.Text) = 0 Then
                                '    'enter Study Number and ID
                                '    var1 = frmSD.dgvSDStudy.Item("CHARSTUDYNUMBER", intRow).Value
                                '    var2 = frmSD.dgvSDStudy.Item("ID_TBLGUWUSTUDIES", intRow).Value
                                '    arrC(4, int1) = var1
                                '    arrC(5, int1) = var2
                                'End If
                            End If
                        End If
                    End If
                Next

                mw = 150

            Case "GroupDetails"
                tblT = tblGuWuPKRoutes
                dtbl = tblRoute

                If frmSD.dgvRoutes.Rows.Count = 0 Then
                    GoTo end1
                End If

                If frmSD.dgvRoutes.CurrentRow Is Nothing Then
                    intRow = 0
                Else
                    intRow = frmSD.dgvRoutes.CurrentRow.Index
                End If

                For Count1 = 0 To intCols - 1
                    str1 = dgv2.Columns(Count1).Name
                    If StrComp(str1, "ID_TBLGUWUPKROUTES", CompareMethod.Text) = 0 Or StrComp(str1, "ID_TBLGUWUSTUDIES", CompareMethod.Text) = 0 Or StrComp(str1, "ID_TBLGUWUASSAY", CompareMethod.Text) = 0 Or StrComp(str1, "ID_TBLGUWUPKGROUPS", CompareMethod.Text) = 0 Or StrComp(str1, "ID_TBLSTUDIES", CompareMethod.Text) = 0 Then
                    Else
                        strID = Mid(str1, 1, 3)
                        'If dgv2.Columns(Count1).Visible Or StrComp(strID, "ID_", CompareMethod.Text) = 0 Then
                        bool = dgv2.Columns(Count1).ReadOnly
                        If bool Then
                        Else
                            int1 = int1 + 1
                            arrC(1, int1) = dgv2.Columns(Count1).Name
                            arrC(2, int1) = dgv2.Columns(Count1).HeaderText
                        End If
                        'End If
                    End If
                Next

                mw = dgv1.Width - (dgv1.RowHeadersWidth * 1.2)
                mw = 150

        End Select
        intRows = int1

        If intRows = 0 Then
            GoTo end1
        End If

        For Count1 = 1 To intRows
            str1 = arrC(1, Count1)
            var1 = tblT.Columns(str1).DataType.ToString
            arrC(3, Count1) = var1
            ''''''''''''console.writeline(var1.ToString)
        Next


        If dtbl.Columns.Count > 0 Then
            GoTo end1
        End If
        'datatypes
        'system.String
        'system.DateTime
        'system.Int16

        'add columns

        Dim col1 As New DataColumn
        col1.ColumnName = "ColumnName"
        dtbl.Columns.Add(col1)

        Dim col2 As New DataColumn
        col2.ColumnName = "ColumnHeader"
        dtbl.Columns.Add(col2)

        Dim col3 As New DataColumn
        col3.ColumnName = "ColumnValue"
        dtbl.Columns.Add(col3)

        Dim col4 As New DataColumn
        col4.ColumnName = "ColumnID"
        dtbl.Columns.Add(col4)

        Dim col5 As New DataColumn
        col5.ColumnName = "ColumnValueActual"
        dtbl.Columns.Add(col5)

        Dim col6 As New DataColumn
        col6.ColumnName = "ColumnDataType"
        dtbl.Columns.Add(col6)

        Dim col7 As New DataColumn
        col7.ColumnName = "ColumnReadOnly"
        dtbl.Columns.Add(col7)

        For Count1 = 1 To intRows
            Dim row As DataRow = dtbl.NewRow
            row("ColumnName") = arrC(1, Count1)
            row("ColumnHeader") = arrC(2, Count1)
            dtbl.Rows.Add(row)
        Next

        Dim dv As system.data.dataview = New DataView(dtbl)
        dv.AllowDelete = False
        dv.AllowNew = False

        dgv1.DataSource = dv

        dgv1.ReadOnly = False

        'configure datatype cells
        For Count1 = 1 To intRows
            str1 = arrC(1, Count1) 'columnname
            str2 = arrC(3, Count1) 'datatype
            str3 = Mid(str1, 1, 2) 'look for dt
            str4 = Mid(str1, 1, 4) 'look for bool

            If StrComp(str3, "DT", CompareMethod.Text) = 0 Then
                'make cell datetime format

            ElseIf StrComp(str4, "BOOL", CompareMethod.Text) = 0 Then
                Select Case strMod
                    Case "Studies"
                        dgv1("ColumnValue", Count1 - 1).Value = "FALSE"
                End Select
            ElseIf StrComp(str1, "ID_TBLGUWUPROJECTS", CompareMethod.Text) = 0 And StrComp(strMod, "Studies", CompareMethod.Text) = 0 Then
                dgv1("ColumnValue", Count1 - 1).Value = arrC(4, Count1)
                dgv1("ColumnID", Count1 - 1).Value = arrC(5, Count1)
            End If

            dgv1("ColumnDataType", Count1 - 1).Value = arrC(3, Count1)

            dgv1("ColumnReadOnly", Count1 - 1).Value = "FALSE" 'start default as false

        Next

        dgv1.Columns("ColumnHeader").HeaderText = "Item"
        dgv1.Columns("ColumnHeader").ReadOnly = True
        dgv1.Columns("ColumnValue").HeaderText = "Value"
        dgv1.Columns("ColumnDataType").HeaderText = "Data Type"

        dgv1.Columns("ColumnHeader").SortMode = DataGridViewColumnSortMode.NotSortable
        dgv1.Columns("ColumnValue").SortMode = DataGridViewColumnSortMode.NotSortable

        Dim intDo As Short
        intDo = 1
        If intDo = 1 Then
            dgv1.Columns("ColumnName").Visible = False
            dgv1.Columns("ColumnValueActual").Visible = False
            dgv1.Columns("ColumnDataType").Visible = False
            dgv1.Columns("ColumnReadOnly").Visible = False
            dgv1.Columns("ColumnID").Visible = False
        Else
            dgv1.Columns("ColumnName").Visible = True
            dgv1.Columns("ColumnValueActual").Visible = True
            dgv1.Columns("ColumnDataType").Visible = True
            dgv1.Columns("ColumnReadOnly").Visible = True
            dgv1.Columns("ColumnID").Visible = True
        End If

        dgv1.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
        dgv1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        dgv1.AllowUserToResizeColumns = True
        dgv1.AllowUserToResizeRows = True
        dgv1.RowHeadersWidth = 25
        dgv1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        dgv1.DefaultCellStyle.WrapMode = DataGridViewTriState.True

        If intDo = 1 Then
            dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            dgv1.Columns.Item("ColumnValue").MinimumWidth = mw
            Try
                dgv1.Columns.Item("ColumnHeader").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                dgv1.Columns.Item("ColummValue").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            Catch ex As Exception

            End Try
        Else
            dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            dgv1.Columns.Item("ColumnValue").MinimumWidth = mw
            Try
                dgv1.Columns.Item("ColumnHeader").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                dgv1.Columns.Item("ColummValue").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            Catch ex As Exception

            End Try
        End If

        dgv1.AutoResizeColumns()

        dgv1.CurrentCell = dgv1.Item("ColumnValue", 0)

end1:

        dgv2.ReadOnly = True

        Call FinalConfigDGV(strMod)

    End Sub

    Sub FinalConfigDGV(ByVal strMod As String)

        Dim dgv As DataGridView
        Dim intCols As Short
        Dim Count1 As Short
        Dim str1 As String
        Dim boolTab As Boolean

        boolTab = False

        Select Case strMod
            Case "Projects"
                If Me.rbProjectTab.Checked Then
                    boolTab = True
                Else
                    boolTab = False
                End If
                If boolSourceSD Then
                    dgv = Me.dgvSDProject
                Else
                    dgv = Me.dgvwProject
                End If
                dgv = Me.dgvSDProject
                intCols = dgv.Columns.Count

                If boolTab Then
                Else
                    For Count1 = 0 To intCols - 1
                        str1 = dgv.Columns(Count1).Name
                        If StrComp(str1, "CHARPROJECTNUM", CompareMethod.Text) = 0 Then
                            dgv.Columns(Count1).Visible = True
                        Else
                            dgv.Columns(Count1).Visible = False
                        End If
                    Next
                End If


            Case "Studies"
                If Me.rbStudyTab.Checked Then
                    boolTab = True
                Else
                    boolTab = False
                End If

                If boolSourceSD Then
                    dgv = Me.dgvSDStudy
                Else
                    dgv = Me.dgvwStudy
                End If
                dgv = Me.dgvSDStudy
                intCols = dgv.Columns.Count

                If boolTab Then
                Else
                    For Count1 = 0 To intCols - 1
                        str1 = dgv.Columns(Count1).Name
                        If StrComp(str1, "CHARSTUDYNUMBER", CompareMethod.Text) = 0 Then
                            dgv.Columns(Count1).Visible = True
                        Else
                            dgv.Columns(Count1).Visible = False
                        End If
                    Next
                End If

            Case "Assays"
                If Me.rbStudyTab.Checked Then
                    boolTab = True
                Else
                    boolTab = False
                End If

                If boolSourceSD Then
                    dgv = Me.dgvAssays
                Else
                    GoTo end1
                    'dgv = Me.dgvwStudy
                End If
                dgv = Me.dgvAssays
                intCols = dgv.Columns.Count

                If boolTab Then
                Else
                    For Count1 = 0 To intCols - 1
                        str1 = dgv.Columns(Count1).Name
                        If StrComp(str1, "CHARASSAYNAME", CompareMethod.Text) = 0 Then
                            dgv.Columns(Count1).Visible = True
                        Else
                            dgv.Columns(Count1).Visible = False
                        End If
                    Next
                End If

        End Select

end1:

    End Sub

    Sub Configdgv1Again(ByVal strMod As String)

        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim boolVis As Boolean
        Dim dgv As DataGridView
        Dim boolTab As Boolean

        Select Case strMod
            Case "Projects"
                dgv = Me.dgvSDProject
                If Me.rbProjectTab.Checked Then
                    boolTab = True
                Else
                    boolTab = False
                End If
                'enter column headertexts
                For Count1 = 0 To dgv.ColumnCount - 1

                    str4 = dgv.Columns(Count1).Name
                    str1 = ""
                    boolVis = False
                    Select Case str4
                        Case "CHARPROJECTNAME"
                            str1 = "CHARPROJECTNAME"
                            str2 = "Project Name"
                            ''boolRO = True
                            If boolTab Then
                                boolVis = True
                            Else
                                boolVis = False
                            End If
                        Case "CHARPROJECTNUM"
                            str1 = "CHARPROJECTNUM"
                            str2 = "Project Number"
                            ''boolRO = True
                            If boolTab Then
                                boolVis = True
                            Else
                                boolVis = True
                            End If

                        Case "CHARPROJECTDESCR"
                            str1 = "CHARPROJECTDESCR"
                            str2 = "Project Description"
                            ''boolRO = True
                            If boolTab Then
                                boolVis = True
                            Else
                                boolVis = False
                            End If
                        Case "CHARUSERIDINIT"
                            str1 = "CHARUSERIDINIT"
                            str2 = "User ID Intialized"
                            'boolRO = True
                            If boolTab Then
                                boolVis = True
                            Else
                                boolVis = False
                            End If
                        Case "CHARUSERNAMEINIT"
                            str1 = "CHARUSERNAMEINIT"
                            str2 = "User Name Initialized"
                            'boolRO = True
                            If boolTab Then
                                boolVis = True
                            Else
                                boolVis = False
                            End If
                        Case "DTINIT"
                            str1 = "DTINIT"
                            str2 = "Date Initialized"
                            'boolRO = True
                            If boolTab Then
                                boolVis = True
                            Else
                                boolVis = False
                            End If
                        Case "CHARUSERIDMOD"
                            str1 = "CHARUSERIDMOD"
                            str2 = "User ID Modified"
                            'boolRO = True
                            If boolTab Then
                                boolVis = True
                            Else
                                boolVis = False
                            End If
                        Case "CHARUSERNAMEMOD"
                            str1 = "CHARUSERNAMEMOD"
                            str2 = "User Name Modified"
                            'boolRO = True
                            If boolTab Then
                                boolVis = True
                            Else
                                boolVis = False
                            End If
                        Case "DTMOD"
                            str1 = "DTMOD"
                            str2 = "Date Modified"
                            'boolRO = True
                            If boolTab Then
                                boolVis = True
                            Else
                                boolVis = False
                            End If
                    End Select
                    If Len(str1) = 0 Then
                    Else
                        'dgv.Columns.Item(str1).Visible = True
                        'dgv.Columns.Item(str1).HeaderText = str2
                        'dgv.Columns.Item(str1).ReadOnly = 'boolRO
                        dgv.Columns.Item(str1).Visible = boolVis
                    End If
                Next
            Case "Studies"
                dgv = Me.dgvSDStudy
                If Me.rbStudyTab.Checked Then
                    boolTab = True
                Else
                    boolTab = False
                End If

                'enter column headertexts
                For Count1 = 0 To dgv.ColumnCount - 1
                    str4 = dgv.Columns(Count1).Name
                    str1 = ""
                    'boolRO = False
                    Select Case str4
                        Case "ID_TBLGUWUPROJECTS"
                            str1 = "ID_TBLGUWUPROJECTS"
                            str2 = "Project Number"
                            'boolRO = True
                            If boolTab Then
                                boolVis = False
                            Else
                                boolVis = False
                            End If

                        Case "ID_TBLCONFIGREPORTTYPE"
                            str1 = "ID_TBLCONFIGREPORTTYPE"
                            str2 = "Study Type"
                            If boolTab Then
                                boolVis = False
                            Else
                                boolVis = False
                            End If

                        Case "ID_TBLGUWUSTUDYSTAT"
                            str1 = "ID_TBLGUWUSTUDYSTAT"
                            str2 = "Study Status"
                            If boolTab Then
                                boolVis = False
                            Else
                                boolVis = False
                            End If

                        Case "ID_TBLGUWUSTUDYDESIGNTYPE"
                            str1 = "ID_TBLGUWUSTUDYDESIGNTYPE"
                            str2 = "Study Design Type"
                            If boolTab Then
                                boolVis = False
                            Else
                                boolVis = False
                            End If

                        Case "CHARSTUDYNAME"
                            str1 = "CHARSTUDYNAME"
                            str2 = "Study Name"
                            'boolRO = False
                            If boolTab Then
                                boolVis = True
                            Else
                                boolVis = False
                            End If

                        Case "CHARSTUDYNUMBER"
                            str1 = "CHARSTUDYNUMBER"
                            str2 = "Study Number"
                            'boolRO = False
                            If boolTab Then
                                boolVis = True
                            Else
                                boolVis = True
                            End If
                        Case "CHARSTUDYDESCR"
                            str1 = "CHARSTUDYDESCR"
                            str2 = "Study Description"
                            'boolRO = False
                            If boolTab Then
                                boolVis = True
                            Else
                                boolVis = False
                            End If
                        Case "BOOLISGLP"
                            str1 = "BOOLISGLP"
                            str2 = "Is GLP"
                            'boolRO = False
                            If boolTab Then
                                boolVis = True
                            Else
                                boolVis = False
                            End If
                        Case "DTSTUDYSTARTPRE"
                            str1 = "DTSTUDYSTARTPRE"
                            str2 = "Study Start Predicted"
                            'boolRO = False
                            If boolTab Then
                                boolVis = True
                            Else
                                boolVis = False
                            End If
                        Case "DTSTUDYSTARTACT"
                            str1 = "DTSTUDYSTARTACT"
                            str2 = "Study Start Actual"
                            'boolRO = False
                            If boolTab Then
                                boolVis = True
                            Else
                                boolVis = False
                            End If
                        Case "DTSTUDYENDPRED"
                            str1 = "DTSTUDYENDPRED"
                            str2 = "Study End Predicted"
                            'boolRO = False
                            If boolTab Then
                                boolVis = True
                            Else
                                boolVis = False
                            End If
                        Case "DTSTUDYENDACT"
                            str1 = "DTSTUDYENDACT"
                            str2 = "Study End Actual"
                            'boolRO = False
                            If boolTab Then
                                boolVis = True
                            Else
                                boolVis = False
                            End If
                        Case "DTEXTRACTIONDATE"
                            str1 = "DTEXTRACTIONDATE"
                            str2 = "Sample Extr. Date"
                            'boolRO = False
                            If boolTab Then
                                boolVis = True
                            Else
                                boolVis = False
                            End If
                        Case "CHARNOTEBOOKREF"
                            str1 = "CHARNOTEBOOKREF"
                            str2 = "Notebook Ref"
                            'boolRO = False
                            If boolTab Then
                                boolVis = True
                            Else
                                boolVis = False
                            End If
                        Case "CHARUSERIDINIT"
                            str1 = "CHARUSERIDINIT"
                            str2 = "User ID Intialized"
                            'boolRO = True
                            If boolTab Then
                                boolVis = True
                            Else
                                boolVis = False
                            End If
                        Case "CHARUSERNAMEINIT"
                            str1 = "CHARUSERNAMEINIT"
                            str2 = "User Name Initialized"
                            'boolRO = True
                            If boolTab Then
                                boolVis = True
                            Else
                                boolVis = False
                            End If
                        Case "DTINIT"
                            str1 = "DTINIT"
                            str2 = "Date Initialized"
                            'boolRO = True
                            If boolTab Then
                                boolVis = True
                            Else
                                boolVis = False
                            End If
                        Case "CHARUSERIDMOD"
                            str1 = "CHARUSERIDMOD"
                            str2 = "User ID Modified"
                            'boolRO = True
                            If boolTab Then
                                boolVis = True
                            Else
                                boolVis = False
                            End If
                        Case "CHARUSERNAMEMOD"
                            str1 = "CHARUSERNAMEMOD"
                            str2 = "User Name Modified"
                            'boolRO = True
                            If boolTab Then
                                boolVis = True
                            Else
                                boolVis = False
                            End If
                        Case "DTMOD"
                            str1 = "DTMOD"
                            str2 = "Date Modified"
                            'boolRO = True
                            If boolTab Then
                                boolVis = True
                            Else
                                boolVis = False
                            End If
                        Case "BOOLINCLUDE"
                            str1 = "BOOLINCLUDE"
                            str2 = "Included"
                            'boolRO = True
                            boolVis = False
                        Case "BOOLACCEPTED"
                            str1 = "BOOLACCEPTED"
                            str2 = "Accepted"
                            'boolRO = True
                            boolVis = False

                    End Select
                    If Len(str1) = 0 Then
                    Else
                        dgv.Columns.Item(str1).Visible = boolVis
                        'dgv.Columns.Item(str1).HeaderText = str2
                        'dgv.Columns.Item(str1).ReadOnly = boolRO
                    End If
                Next

            Case "Assays"
                'dgv = Me.dgvAssays
                'If Me.rbStudyTab.Checked Then
                '    boolTab = True
                'Else
                '    boolTab = False
                'End If

                ''enter column headertexts
                'For Count1 = 0 To dgv.ColumnCount - 1
                '    str4 = dgv.Columns(Count1).Name
                '    str1 = ""
                '    'boolRO = False
                '    Select Case str4
                '        Case "ID_TBLGUWUPROJECTS"
                '            str1 = "ID_TBLGUWUPROJECTS"
                '            str2 = "Project Number"
                '            'boolRO = True
                '            If boolTab Then
                '                boolVis = False
                '            Else
                '                boolVis = False
                '            End If

                '        Case "ID_TBLCONFIGREPORTTYPE"
                '            str1 = "ID_TBLCONFIGREPORTTYPE"
                '            str2 = "Study Type"
                '            If boolTab Then
                '                boolVis = False
                '            Else
                '                boolVis = False
                '            End If

                '        Case "ID_TBLGUWUSTUDYSTAT"
                '            str1 = "ID_TBLGUWUSTUDYSTAT"
                '            str2 = "Study Status"
                '            If boolTab Then
                '                boolVis = False
                '            Else
                '                boolVis = False
                '            End If

                '        Case "ID_TBLGUWUSTUDYDESIGNTYPE"
                '            str1 = "ID_TBLGUWUSTUDYDESIGNTYPE"
                '            str2 = "Study Design Type"
                '            If boolTab Then
                '                boolVis = False
                '            Else
                '                boolVis = False
                '            End If

                '        Case "CHARSTUDYNAME"
                '            str1 = "CHARSTUDYNAME"
                '            str2 = "Study Name"
                '            'boolRO = False
                '            If boolTab Then
                '                boolVis = True
                '            Else
                '                boolVis = False
                '            End If

                '        Case "CHARSTUDYNUMBER"
                '            str1 = "CHARSTUDYNUMBER"
                '            str2 = "Study Number"
                '            'boolRO = False
                '            If boolTab Then
                '                boolVis = True
                '            Else
                '                boolVis = True
                '            End If
                '        Case "CHARSTUDYDESCR"
                '            str1 = "CHARSTUDYDESCR"
                '            str2 = "Study Description"
                '            'boolRO = False
                '            If boolTab Then
                '                boolVis = True
                '            Else
                '                boolVis = False
                '            End If
                '        Case "BOOLISGLP"
                '            str1 = "BOOLISGLP"
                '            str2 = "Is GLP"
                '            'boolRO = False
                '            If boolTab Then
                '                boolVis = True
                '            Else
                '                boolVis = False
                '            End If
                '        Case "DTSTUDYSTARTPRE"
                '            str1 = "DTSTUDYSTARTPRE"
                '            str2 = "Study Start Predicted"
                '            'boolRO = False
                '            If boolTab Then
                '                boolVis = True
                '            Else
                '                boolVis = False
                '            End If
                '        Case "DTSTUDYSTARTACT"
                '            str1 = "DTSTUDYSTARTACT"
                '            str2 = "Study Start Actual"
                '            'boolRO = False
                '            If boolTab Then
                '                boolVis = True
                '            Else
                '                boolVis = False
                '            End If
                '        Case "DTSTUDYENDPRED"
                '            str1 = "DTSTUDYENDPRED"
                '            str2 = "Study End Predicted"
                '            'boolRO = False
                '            If boolTab Then
                '                boolVis = True
                '            Else
                '                boolVis = False
                '            End If
                '        Case "DTSTUDYENDACT"
                '            str1 = "DTSTUDYENDACT"
                '            str2 = "Study End Actual"
                '            'boolRO = False
                '            If boolTab Then
                '                boolVis = True
                '            Else
                '                boolVis = False
                '            End If
                '        Case "DTEXTRACTIONDATE"
                '            str1 = "DTEXTRACTIONDATE"
                '            str2 = "Sample Extr. Date"
                '            'boolRO = False
                '            If boolTab Then
                '                boolVis = True
                '            Else
                '                boolVis = False
                '            End If
                '        Case "CHARNOTEBOOKREF"
                '            str1 = "CHARNOTEBOOKREF"
                '            str2 = "Notebook Ref"
                '            'boolRO = False
                '            If boolTab Then
                '                boolVis = True
                '            Else
                '                boolVis = False
                '            End If
                '        Case "CHARUSERIDINIT"
                '            str1 = "CHARUSERIDINIT"
                '            str2 = "User ID Intialized"
                '            'boolRO = True
                '            If boolTab Then
                '                boolVis = True
                '            Else
                '                boolVis = False
                '            End If
                '        Case "CHARUSERNAMEINIT"
                '            str1 = "CHARUSERNAMEINIT"
                '            str2 = "User Name Initialized"
                '            'boolRO = True
                '            If boolTab Then
                '                boolVis = True
                '            Else
                '                boolVis = False
                '            End If
                '        Case "DTINIT"
                '            str1 = "DTINIT"
                '            str2 = "Date Initialized"
                '            'boolRO = True
                '            If boolTab Then
                '                boolVis = True
                '            Else
                '                boolVis = False
                '            End If
                '        Case "CHARUSERIDMOD"
                '            str1 = "CHARUSERIDMOD"
                '            str2 = "User ID Modified"
                '            'boolRO = True
                '            If boolTab Then
                '                boolVis = True
                '            Else
                '                boolVis = False
                '            End If
                '        Case "CHARUSERNAMEMOD"
                '            str1 = "CHARUSERNAMEMOD"
                '            str2 = "User Name Modified"
                '            'boolRO = True
                '            If boolTab Then
                '                boolVis = True
                '            Else
                '                boolVis = False
                '            End If
                '        Case "DTMOD"
                '            str1 = "DTMOD"
                '            str2 = "Date Modified"
                '            'boolRO = True
                '            If boolTab Then
                '                boolVis = True
                '            Else
                '                boolVis = False
                '            End If
                '    End Select
                '    If Len(str1) = 0 Then
                '    Else
                '        dgv.Columns.Item(str1).Visible = boolVis
                '        'dgv.Columns.Item(str1).HeaderText = str2
                '        'dgv.Columns.Item(str1).ReadOnly = boolRO
                '    End If
                'Next

        End Select

        If boolTab Then
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgv.AutoResizeColumns()
        Else
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        End If

    End Sub

    Sub ConfigureDGVs()

        Call Configure_dgvProject()

        Call Configure_dgvStudy()

        Call Configure_dgvAssay()

        Call Configure_dgvGroup()

        Call Configure_dgvRoute()

        Call Configure_dgvTimePoint()

        Call Configure_dgvPatient()

        Call Configure_dgvCmpd()

        Call Configure_dgvPI()

        Call Configure_dgvAnalyst()


    End Sub

    Sub Configure_dgvProject()
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim Count1 As Short
        Dim cw As Short

        For Count1 = 1 To 4
            Select Case Count1
                Case 1
                    dgv = Me.dgvSDProject
                    cw = 25
                Case 2
                    dgv = Me.dgvwProject
                    cw = 25
                Case 3
                    dgv = Me.dgvSDProjectS
                    cw = 10
                Case 4
                    dgv = Me.dgvwProjectS
                    cw = 10
            End Select
            dgv.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgv.AllowUserToResizeColumns = True
            dgv.AllowUserToResizeRows = True
            dgv.RowHeadersWidth = cw '25
            dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True

            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            Select Case Count1
                Case 1
                    Call ConfigProjectTableSD()
                Case 2
                    Call ConfigProjectTableW()
            End Select
        Next

    End Sub

    Sub ConfigProjectTableSD()

        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        Dim boolRO As Boolean
        'Dim rs As New ADODB.Recordset
        Dim int1 As Short
        Dim intCol As Short
        Dim boolVis As Boolean

        'dv = tblwSTUDY.DefaultView
        dv = New DataView(tblGuWuProjects)
        dv.AllowDelete = False
        dv.AllowEdit = False
        dv.AllowNew = False

        dgv = Me.dgvSDProject

        dgv.DataSource = dv

        'hide all columns
        For Count1 = 0 To dgv.Columns.Count - 1
            dgv.Columns.Item(Count1).Visible = False
            'dgv.Columns.Item(Count1).DisplayIndex = dgv.Columns.Count - 1
            'dgv.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        'enter column headertexts
        For Count1 = 0 To dgv.ColumnCount - 1
            str4 = dgv.Columns(Count1).Name
            str1 = ""
            Select Case str4
                Case "CHARPROJECTNAME"
                    str1 = "CHARPROJECTNAME"
                    str2 = "Project Name"
                    boolRO = False
                    intCol = Count1
                Case "CHARPROJECTNUM"
                    str1 = "CHARPROJECTNUM"
                    str2 = "Project Number"
                    boolRO = False
                Case "CHARPROJECTDESCR"
                    str1 = "CHARPROJECTDESCR"
                    str2 = "Project Description"
                    boolRO = False
                Case "CHARUSERIDINIT"
                    str1 = "CHARUSERIDINIT"
                    str2 = "User ID Intialized"
                    boolRO = True
                Case "CHARUSERNAMEINIT"
                    str1 = "CHARUSERNAMEINIT"
                    str2 = "User Name Initialized"
                    boolRO = True
                Case "DTINIT"
                    str1 = "DTINIT"
                    str2 = "Date Initialized"
                    boolRO = True
                Case "CHARUSERIDMOD"
                    str1 = "CHARUSERIDMOD"
                    str2 = "User ID Modified"
                    boolRO = True
                Case "CHARUSERNAMEMOD"
                    str1 = "CHARUSERNAMEMOD"
                    str2 = "User Name Modified"
                    boolRO = True
                Case "DTMOD"
                    str1 = "DTMOD"
                    str2 = "Date Modified"
                    boolRO = True
                Case "BOOLINCLUDE"
                    str1 = "BOOLINCLUDE"
                    str2 = "Included"
                    boolRO = True
                Case "BOOLACCEPTED"
                    str1 = "BOOLACCEPTED"
                    str2 = "Accepted"
                    boolRO = True

            End Select
            If Len(str1) = 0 Then
            Else
                dgv.Columns.Item(str1).Visible = True
                dgv.Columns.Item(str1).HeaderText = str2
                dgv.Columns.Item(str1).ReadOnly = boolRO
            End If
        Next

        'set first row as current row
        dgv.CurrentCell = dgv.Rows.Item(0).Cells(intCol)
        dgv.AutoResizeColumns()


        'now do dgvDSProjectsS
        dgv = Me.dgvSDProjectS
        dgv.DataSource = dv

        'hide all columns
        For Count1 = 0 To dgv.Columns.Count - 1
            dgv.Columns.Item(Count1).Visible = False
            'dgv.Columns.Item(Count1).DisplayIndex = dgv.Columns.Count - 1
            'dgv.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        'enter column headertexts
        For Count1 = 0 To dgv.ColumnCount - 1
            str4 = dgv.Columns(Count1).Name
            str1 = ""
            boolVis = False
            Select Case str4
                Case "CHARPROJECTNAME"
                    str1 = "CHARPROJECTNAME"
                    str2 = "Project Name"
                    boolRO = True
                    boolVis = False
                Case "CHARPROJECTNUM"
                    str1 = "CHARPROJECTNUM"
                    str2 = "Project Number"
                    boolRO = True
                    boolVis = True
                Case "CHARPROJECTDESCR"
                    str1 = "CHARPROJECTDESCR"
                    str2 = "Project Description"
                    boolRO = True
                    boolVis = False

            End Select
            If boolVis Then
                intCol = Count1
            End If
            If Len(str1) = 0 Then
            Else
                dgv.Columns.Item(str1).Visible = True
                dgv.Columns.Item(str1).HeaderText = str2
                dgv.Columns.Item(str1).ReadOnly = boolRO
                dgv.Columns.Item(str1).Visible = boolVis
            End If
        Next

        'make studytitle column fit to grid
        'If boolW Then
        '    dgv.Columns.Item("StudyTitle").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        '    dgv.Columns.Item("StudyTitle").DefaultCellStyle.WrapMode = DataGridViewTriState.True
        '    dgv.AutoResizeColumns()
        '    dgv.AutoResizeRows()
        'End If

        'set first row as current row
        dgv.CurrentCell = dgv.Rows.Item(0).Cells(intCol)
        dgv.AutoResizeColumns()

    End Sub

    Sub dgvProjectsReadOnly()
        Dim dgv As DataGridView
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim boolRO As Boolean
        Dim intCol As Short
        Dim intRow As Short

        dgv = Me.dgvSDProject

        'enter column headertexts
        For Count1 = 0 To dgv.ColumnCount - 1
            str4 = dgv.Columns(Count1).Name
            str1 = ""
            Select Case str4
                Case "CHARPROJECTNAME"
                    str1 = "CHARPROJECTNAME"
                    str2 = "Project Name"
                    boolRO = False
                    intCol = Count1
                Case "CHARPROJECTNUM"
                    str1 = "CHARPROJECTNUM"
                    str2 = "Project Number"
                    boolRO = False
                Case "CHARPROJECTDESCR"
                    str1 = "CHARPROJECTDESCR"
                    str2 = "Project Description"
                    boolRO = False
                Case "CHARUSERIDINIT"
                    str1 = "CHARUSERIDINIT"
                    str2 = "User ID Intialized"
                    boolRO = True
                Case "CHARUSERNAMEINIT"
                    str1 = "CHARUSERNAMEINIT"
                    str2 = "User Name Initialized"
                    boolRO = True
                Case "DTINIT"
                    str1 = "DTINIT"
                    str2 = "Date Initialized"
                    boolRO = True
                Case "CHARUSERIDMOD"
                    str1 = "CHARUSERIDMOD"
                    str2 = "User ID Modified"
                    boolRO = True
                Case "CHARUSERNAMEMOD"
                    str1 = "CHARUSERNAMEMOD"
                    str2 = "User Name Modified"
                    boolRO = True
                Case "DTMOD"
                    str1 = "DTMOD"
                    str2 = "Date Modified"
                    boolRO = True
                Case "BOOLINCLUDE"
                    str1 = "BOOLINCLUDE"
                    str2 = "Included"
                    boolRO = True
                Case "BOOLACCEPTED"
                    str1 = "BOOLACCEPTED"
                    str2 = "Accepted"
                    boolRO = True
            End Select
            If Len(str1) = 0 Then
            Else
                'dgv.Columns.Item(str1).ReadOnly = boolRO
                intRow = FindRow(Me.dgvProj, str1)
                If intRow = -1 Then
                Else
                    Me.dgvProj.Rows(intRow).Cells("ColumnValue").ReadOnly = boolRO
                End If
            End If
        Next

    End Sub

    Function FindRow(ByVal dgv As DataGridView, ByVal strName As String) As Short

        Dim intRows As Short
        Dim Count1 As Short
        Dim str1 As String

        intRows = dgv.Rows.Count
        FindRow = -1
        For Count1 = 0 To intRows - 1
            str1 = dgv("ColumnName", Count1).Value
            If StrComp(str1, strName, CompareMethod.Text) = 0 Then
                FindRow = Count1
                Exit For
            End If
        Next


    End Function

    Sub ConfigProjectTableW()

    End Sub

    Sub Configure_dgvAnalyst()
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim Count1 As Short
        Dim cw

        If boolSourceSD Then
            dgv = Me.dgvAnalyst
        Else
            GoTo end1
        End If

        For Count1 = 1 To 1
            Select Case Count1
                Case 1
                    dgv = Me.dgvAnalyst
                    cw = 10


            End Select
            dgv.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgv.AllowUserToResizeColumns = True
            dgv.AllowUserToResizeRows = True
            dgv.RowHeadersWidth = cw '25
            dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True

            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            'this done already in PI
            'Select Case Count1
            '    Case 1
            '        Call ConfigPersTableSD(True)

            'End Select
        Next

end1:

    End Sub

    Sub Configure_dgvPI()
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim Count1 As Short
        Dim cw

        If boolSourceSD Then
            dgv = Me.dgvPI
        Else
            GoTo end1
        End If

        For Count1 = 1 To 1
            Select Case Count1
                Case 1
                    dgv = Me.dgvPI
                    cw = 10


            End Select
            dgv.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgv.AllowUserToResizeColumns = True
            dgv.AllowUserToResizeRows = True
            dgv.RowHeadersWidth = cw '25
            dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True

            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            Select Case Count1
                Case 1
                    Call ConfigPersTableSD(True)

            End Select
        Next

end1:

    End Sub

    Sub Configure_dgvCmpd()
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim Count1 As Short
        Dim cw

        If boolSourceSD Then
            dgv = Me.dgvCmpd
        Else
            GoTo end1
        End If

        For Count1 = 1 To 1
            Select Case Count1
                Case 1
                    dgv = Me.dgvCmpd
                    cw = 10


            End Select
            dgv.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgv.AllowUserToResizeColumns = True
            dgv.AllowUserToResizeRows = True
            dgv.RowHeadersWidth = cw '25
            dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True

            'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            Select Case Count1
                Case 1
                    Call ConfigCmpdTableSD(True)

            End Select
        Next

end1:

    End Sub


    Sub Configure_dgvCmpdLot()
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim Count1 As Short
        Dim cw

        If boolSourceSD Then
            dgv = Me.dgvLotNum
        Else
            GoTo end1
        End If

        For Count1 = 1 To 1
            Select Case Count1
                Case 1
                    dgv = Me.dgvLotNum
                    cw = 10


            End Select
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgv.AllowUserToResizeColumns = True
            dgv.AllowUserToResizeRows = True
            dgv.RowHeadersWidth = cw '25
            dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True

            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            Select Case Count1
                Case 1
                    'Call ConfigCmpdTableSD(True)


            End Select
        Next

end1:

    End Sub

    Sub Configure_dgvPatient()
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim Count1 As Short
        Dim cw

        dgv = Me.dgvPatients
        If boolSourceSD Then
            dgv = Me.dgvPatients
        Else
            GoTo end1
        End If

        For Count1 = 1 To 1
            Select Case Count1
                Case 1
                    dgv = Me.dgvPatients
                    cw = 10


            End Select
            dgv.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgv.AllowUserToResizeColumns = True
            dgv.AllowUserToResizeRows = True
            dgv.RowHeadersWidth = cw '25
            dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True

            dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            'dgv.Columns("NUMTIMEPOINT").SortMode = DataGridViewColumnSortMode.NotSortable
            Try
                dgv.Columns("CHARSUBJECTNAME").SortMode = DataGridViewColumnSortMode.NotSortable
            Catch ex As Exception
                Dim str1 As String
                str1 = ""
            End Try

            Select Case Count1
                Case 1
                    Call ConfigSubjectTableSD(True)

            End Select
        Next



end1:

    End Sub

    Sub Configure_dgvTimePoint()
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim Count1 As Short
        Dim cw

        If boolSourceSD Then
            dgv = Me.dgvGroupTimePoints
        Else
            GoTo end1
        End If

        For Count1 = 1 To 1
            Select Case Count1
                Case 1
                    dgv = Me.dgvGroupTimePoints
                    cw = 10


            End Select
            dgv.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgv.AllowUserToResizeColumns = True
            dgv.AllowUserToResizeRows = True
            dgv.RowHeadersWidth = cw '25
            dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True

            dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            'dgv.Columns("NUMTIMEPOINT").SortMode = DataGridViewColumnSortMode.NotSortable
            Try
                dgv.Columns("NUMTIMEPOINT").SortMode = DataGridViewColumnSortMode.NotSortable
            Catch ex As Exception
                Dim str1 As String
                str1 = ""
            End Try

            Select Case Count1
                Case 1
                    Call ConfigTimePointTableSD(True)

            End Select
        Next



end1:

    End Sub

    Sub Configure_dgvRoute()
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim Count1 As Short
        Dim cw

        If boolSourceSD Then
            dgv = Me.dgvRoutes
        Else
            GoTo end1
        End If

        For Count1 = 1 To 1
            Select Case Count1
                Case 1
                    dgv = Me.dgvRoutes
                    cw = 10


            End Select
            dgv.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgv.AllowUserToResizeColumns = True
            dgv.AllowUserToResizeRows = True
            dgv.RowHeadersWidth = cw '25
            dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True

            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            Select Case Count1
                Case 1
                    Call ConfigRouteTableSD(True)

            End Select
        Next

end1:

    End Sub

    Sub Configure_dgvGroup()
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim Count1 As Short
        Dim cw

        If boolSourceSD Then
            dgv = Me.dgvGroups
        Else
            GoTo end1
        End If

        For Count1 = 1 To 1
            Select Case Count1
                Case 1
                    dgv = Me.dgvGroups
                    cw = 10


            End Select
            dgv.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgv.AllowUserToResizeColumns = True
            dgv.AllowUserToResizeRows = True
            dgv.RowHeadersWidth = cw '25
            dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True

            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            Select Case Count1
                Case 1
                    Call ConfigGroupTableSD(True)

            End Select
        Next

end1:

    End Sub

    Sub Configure_dgvAssay()
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim Count1 As Short
        Dim cw

        If boolSourceSD Then
            dgv = Me.dgvAssays
        Else
            GoTo end1
        End If

        For Count1 = 1 To 2
            Select Case Count1
                Case 1
                    dgv = Me.dgvAssays
                    cw = 10
                Case 2
                    dgv = Me.dgvAss
                    cw = 10

            End Select
            dgv.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgv.AllowUserToResizeColumns = True
            dgv.AllowUserToResizeRows = True
            dgv.RowHeadersWidth = cw '25
            dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True

            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            Select Case Count1
                Case 1
                    Call ConfigAssayTableSD(True)

            End Select
        Next

end1:

    End Sub

    Sub Configure_dgvStudy()
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim Count1 As Short
        Dim cw

        If boolSourceSD Then
            dgv = Me.dgvSDStudy
        Else
            dgv = Me.dgvwStudy
        End If

        For Count1 = 1 To 2
            Select Case Count1
                Case 1
                    dgv = Me.dgvSDStudy
                    cw = 10
                Case 2
                    dgv = Me.dgvwStudy
                    cw = 10
            End Select
            dgv.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgv.AllowUserToResizeColumns = True
            dgv.AllowUserToResizeRows = True
            dgv.RowHeadersWidth = cw '25
            dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True

            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            Select Case Count1
                Case 1
                    Call ConfigStudyTableSD(True)
                Case 2
                    Call ConfigStudyTableW(True)
            End Select
        Next

    End Sub

    Sub ConfigStudyTableW(ByVal boolW As Boolean)

        Dim dgv As DataGridView
        Dim dv As System.Data.DataView
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim strSQL As String
        Dim boolRO As Boolean
        'Dim rs As New ADODB.Recordset
        Dim int1 As Short
        Dim intCol As Short
        Dim intRows As Short

        'dv = tblwSTUDY.DefaultView
        dv = New DataView(tblwSTUDY)
        dv.AllowDelete = False
        dv.AllowEdit = False
        dv.AllowNew = False

        dgv = Me.dgvwStudy
        intRows = tblwSTUDY.Rows.Count

        dgv.DataSource = dv

        'hide all columns
        For Count1 = 0 To dgv.Columns.Count - 1
            dgv.Columns.Item(Count1).Visible = False
            dgv.Columns.Item(Count1).DisplayIndex = dgv.Columns.Count - 1
            dgv.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        If intRows = 0 Then
        Else
            'enter column headertexts
            For Count1 = 1 To dgv.ColumnCount - 1
                Dim dgc1 As New DataGridTextBoxColumn
                Select Case Count1
                    Case 1
                        str1 = "PROJECTIDTEXT"
                        str2 = "Project ID"
                        boolRO = True
                        intCol = Count1
                    Case 2
                        str1 = "StudyName"
                        str2 = "Study Name"
                        boolRO = True
                    Case 3
                        str1 = "StudyNumber"
                        str2 = "Study #"
                        boolRO = True
                    Case 4
                        str1 = "Species"
                        str2 = "Species"
                        boolRO = True
                    Case 5
                        str1 = "StudyTitle"
                        str2 = "Study Title"
                        boolRO = True
                        'Case 5
                        '    str1 = "PROJECTID"
                        '    str2 = "Project ID"
                        '    boolRO = True
                        'Case 6
                        '    str1 = "STUDYID"
                        '    str2 = "Study ID"
                        '    boolRO = True
                End Select
                dgv.Columns.Item(str1).Visible = True
                dgv.Columns.Item(str1).HeaderText = str2
                dgv.Columns.Item(str1).DisplayIndex = Count1 - 1
            Next

            'make studytitle column fit to grid
            If boolW Then
                dgv.Columns.Item("StudyTitle").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                dgv.Columns.Item("StudyTitle").DefaultCellStyle.WrapMode = DataGridViewTriState.True
                dgv.AutoResizeColumns()
                dgv.AutoResizeRows()
            End If

            'set first row as current row
            'NOT ANY MORE!
            'dgv.CurrentCell = dgv.Rows.Item(0).Cells(intCol)
        End If

    End Sub

    Sub ConfigStudyTableSDdv()
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim id As Long
        Dim intRow As Short
        Dim strF As String
        Dim strS As String


        'filter dv for chosen project
        If Me.dgvSDProjectS.RowCount = 0 Then
            GoTo end1
        ElseIf Me.dgvSDProjectS.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = Me.dgvSDProjectS.CurrentRow.Index
        End If
        id = Me.dgvSDProjectS.Item("ID_TBLGUWUPROJECTS", intRow).Value
        strF = "ID_TBLGUWUPROJECTS = " & id
        strS = "CHARSTUDYNAME ASC"

        dv = New DataView(tblGuWuStudies, strF, strS, DataViewRowState.CurrentRows)
        dv.AllowDelete = False
        dv.AllowEdit = False
        dv.AllowNew = False

        dgv = Me.dgvSDStudy

        dgv.DataSource = dv

end1:

    End Sub

    Sub ConfigStudyTableSD(ByVal boolW As Boolean)
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        Dim boolRO As Boolean
        'Dim rs As New ADODB.Recordset
        Dim int1 As Short
        Dim intCol As Short
        Dim boolVis As Boolean

        If Me.dgvSDProjectS.RowCount = 0 Then
            GoTo end1
        End If

        Call ConfigStudyTableSDdv()

        dgv = Me.dgvSDStudy
        'hide all columns
        For Count1 = 0 To dgv.Columns.Count - 1
            dgv.Columns.Item(Count1).Visible = False
            'dgv.Columns.Item(Count1).DisplayIndex = dgv.Columns.Count - 1
            'dgv.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        'enter column headertexts
        For Count1 = 0 To dgv.ColumnCount - 1
            str4 = dgv.Columns(Count1).Name
            str1 = ""
            boolRO = False
            Select Case str4
                Case "ID_TBLGUWUPROJECTS"
                    str1 = "ID_TBLGUWUPROJECTS"
                    str2 = "Project Number"
                    boolRO = True

                Case "ID_TBLCONFIGREPORTTYPE"
                    str1 = "ID_TBLCONFIGREPORTTYPE"
                    str2 = "Study Type"

                Case "ID_TBLGUWUSTUDYSTAT"
                    str1 = "ID_TBLGUWUSTUDYSTAT"
                    str2 = "Study Status"

                Case "ID_TBLGUWUSTUDYDESIGNTYPE"
                    str1 = "ID_TBLGUWUSTUDYDESIGNTYPE"
                    str2 = "Study Design Type"

                Case "CHARSTUDYNAME"
                    str1 = "CHARSTUDYNAME"
                    str2 = "Study Name"
                    boolRO = False
                    intCol = Count1
                Case "CHARSTUDYNUMBER"
                    str1 = "CHARSTUDYNUMBER"
                    str2 = "Study Number"
                    boolRO = False
                Case "CHARSTUDYDESCR"
                    str1 = "CHARSTUDYDESCR"
                    str2 = "Study Description"
                    boolRO = False
                Case "BOOLISGLP"
                    str1 = "BOOLISGLP"
                    str2 = "Is GLP"
                    boolRO = False
                Case "DTSTUDYSTARTPRE"
                    str1 = "DTSTUDYSTARTPRE"
                    str2 = "Study Start Predicted"
                    boolRO = False
                Case "DTSTUDYSTARTACT"
                    str1 = "DTSTUDYSTARTACT"
                    str2 = "Study Start Actual"
                    boolRO = False
                Case "DTSTUDYENDPRED"
                    str1 = "DTSTUDYENDPRED"
                    str2 = "Study End Predicted"
                    boolRO = False
                Case "DTSTUDYENDACT"
                    str1 = "DTSTUDYENDACT"
                    str2 = "Study End Actual"
                    boolRO = False
                Case "DTEXTRACTIONDATE"
                    str1 = "DTEXTRACTIONDATE"
                    str2 = "Sample Extr. Date"
                    boolRO = False
                Case "CHARNOTEBOOKREF"
                    str1 = "CHARNOTEBOOKREF"
                    str2 = "Notebook Ref"
                    boolRO = False
                Case "CHARUSERIDINIT"
                    str1 = "CHARUSERIDINIT"
                    str2 = "User ID Intialized"
                    boolRO = True
                Case "CHARUSERNAMEINIT"
                    str1 = "CHARUSERNAMEINIT"
                    str2 = "User Name Initialized"
                    boolRO = True
                Case "DTINIT"
                    str1 = "DTINIT"
                    str2 = "Date Initialized"
                    boolRO = True
                Case "CHARUSERIDMOD"
                    str1 = "CHARUSERIDMOD"
                    str2 = "User ID Modified"
                    boolRO = True
                Case "CHARUSERNAMEMOD"
                    str1 = "CHARUSERNAMEMOD"
                    str2 = "User Name Modified"
                    boolRO = True
                Case "DTMOD"
                    str1 = "DTMOD"
                    str2 = "Date Modified"
                    boolRO = True
                Case "BOOLINCLUDE"
                    str1 = "BOOLINCLUDE"
                    str2 = "Included"
                    boolRO = True
                Case "BOOLACCEPTED"
                    str1 = "BOOLACCEPTED"
                    str2 = "Accepted"
                    boolRO = True

            End Select
            If Len(str1) = 0 Then
            Else
                dgv.Columns.Item(str1).Visible = True
                dgv.Columns.Item(str1).HeaderText = str2
                dgv.Columns.Item(str1).ReadOnly = boolRO
            End If
        Next

        'make studytitle column fit to grid
        'If boolW Then
        '    dgv.Columns.Item("StudyTitle").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        '    dgv.Columns.Item("StudyTitle").DefaultCellStyle.WrapMode = DataGridViewTriState.True
        '    dgv.AutoResizeColumns()
        '    dgv.AutoResizeRows()
        'End If


        Try
            'set first row as current row
            dgv.CurrentCell = dgv.Rows.Item(0).Cells(intCol)
            dgv.AutoResizeColumns()
        Catch ex As Exception

        End Try

end1:

    End Sub


    Sub dgvStudiesReadOnly()
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        Dim boolRO As Boolean
        'Dim rs As New ADODB.Recordset
        Dim int1 As Short
        Dim intCol As Short
        Dim boolVis As Boolean
        Dim intRow As Short

        If Me.dgvSDProjectS.RowCount = 0 Then
            GoTo end1
        End If

        dgv = Me.dgvSDStudy

        'enter column readonly
        For Count1 = 0 To dgv.ColumnCount - 1
            str4 = dgv.Columns(Count1).Name
            str1 = ""
            boolRO = False
            Select Case str4

                Case "ID_TBLGUWUPROJECTS"
                    str1 = "ID_TBLGUWUPROJECTS"
                    str2 = "Project Number"
                    boolRO = False

                Case "ID_TBLCONFIGREPORTTYPE"
                    str1 = "ID_TBLCONFIGREPORTTYPE"
                    str2 = "Study Type"

                Case "ID_TBLGUWUSTUDYSTAT"
                    str1 = "ID_TBLGUWUSTUDYSTAT"
                    str2 = "Study Status"

                Case "ID_TBLGUWUSTUDYDESIGNTYPE"
                    str1 = "ID_TBLGUWUSTUDYDESIGNTYPE"
                    str2 = "Study Design Type"

                Case "CHARSTUDYNAME"
                    str1 = "CHARSTUDYNAME"
                    str2 = "Study Name"
                    boolRO = False
                    intCol = Count1
                Case "CHARSTUDYNUMBER"
                    str1 = "CHARSTUDYNUMBER"
                    str2 = "Study Number"
                    boolRO = False
                Case "CHARSTUDYDESCR"
                    str1 = "CHARSTUDYDESCR"
                    str2 = "Study Description"
                    boolRO = False
                Case "BOOLISGLP"
                    str1 = "BOOLISGLP"
                    str2 = "Is GLP"
                    boolRO = False
                Case "DTSTUDYSTARTPRE"
                    str1 = "DTSTUDYSTARTPRE"
                    str2 = "Study Start Predicted"
                    boolRO = False
                Case "DTSTUDYSTARTACT"
                    str1 = "DTSTUDYSTARTACT"
                    str2 = "Study Start Actual"
                    boolRO = False
                Case "DTSTUDYENDPRED"
                    str1 = "DTSTUDYENDPRED"
                    str2 = "Study End Predicted"
                    boolRO = False
                Case "DTSTUDYENDACT"
                    str1 = "DTSTUDYENDACT"
                    str2 = "Study End Actual"
                    boolRO = False
                Case "DTEXTRACTIONDATE"
                    str1 = "DTEXTRACTIONDATE"
                    str2 = "Sample Extr. Date"
                    boolRO = False
                Case "CHARNOTEBOOKREF"
                    str1 = "CHARNOTEBOOKREF"
                    str2 = "Notebook Ref"
                    boolRO = False
                Case "CHARUSERIDINIT"
                    str1 = "CHARUSERIDINIT"
                    str2 = "User ID Intialized"
                    boolRO = True
                Case "CHARUSERNAMEINIT"
                    str1 = "CHARUSERNAMEINIT"
                    str2 = "User Name Initialized"
                    boolRO = True
                Case "DTINIT"
                    str1 = "DTINIT"
                    str2 = "Date Initialized"
                    boolRO = True
                Case "CHARUSERIDMOD"
                    str1 = "CHARUSERIDMOD"
                    str2 = "User ID Modified"
                    boolRO = True
                Case "CHARUSERNAMEMOD"
                    str1 = "CHARUSERNAMEMOD"
                    str2 = "User Name Modified"
                    boolRO = True
                Case "DTMOD"
                    str1 = "DTMOD"
                    str2 = "Date Modified"
                    boolRO = True
                Case "BOOLINCLUDE"
                    str1 = "BOOLINCLUDE"
                    str2 = "Included"
                    boolRO = True
                Case "BOOLACCEPTED"
                    str1 = "BOOLACCEPTED"
                    str2 = "Accepted"
                    boolRO = True

            End Select
            If Len(str1) = 0 Then
            Else
                'dgv.Columns.Item(str1).ReadOnly = boolRO
                intRow = FindRow(Me.dgvStud, str1)
                If intRow = -1 Then
                Else
                    Me.dgvStud.Rows(intRow).Cells("ColumnValue").ReadOnly = boolRO
                End If
            End If
        Next

end1:

    End Sub

    Sub dgvAssayReadOnly()
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        Dim boolRO As Boolean
        'Dim rs As New ADODB.Recordset
        Dim int1 As Short
        Dim intCol As Short
        Dim boolVis As Boolean
        Dim intRow As Short

        If Me.dgvAssays.RowCount = 0 Then
            GoTo end1
        End If

        dgv = Me.dgvAssays

        'enter column readonly
        For Count1 = 0 To dgv.ColumnCount - 1
            str4 = dgv.Columns(Count1).Name
            str1 = ""
            boolRO = False
            Select Case str4

                Case "ID_TBLGUWUASSAY"
                    str1 = "ID_TBLGUWUASSAY"
                    str2 = "ID_TBLGUWUASSAY"
                    boolRO = True
                    boolVis = False
                Case "ID_TBLGUWUSTUDIES"
                    str1 = "ID_TBLGUWUSTUDIES"
                    str2 = "ID_TBLGUWUSTUDIES"
                    boolRO = True
                    boolVis = False
                Case "ID_TBLSTUDIES"
                    str1 = "ID_TBLSTUDIES"
                    str2 = "ID_TBLSTUDIES"
                    boolRO = True
                    boolVis = False
                Case "CHARASSAYNAME"
                    str1 = "CHARASSAYNAME"
                    str2 = "Assay Name"
                    boolRO = False
                    boolVis = True
                Case "DTASSAYDATE"
                    str1 = "DTASSAYDATE"
                    str2 = "Assay Date"
                    boolRO = False
                    boolVis = True
                Case "ID_TBLGUWUSPECIES"
                    str1 = "ID_TBLGUWUSPECIES"
                    str2 = "ID_TBLGUWUSPECIES"
                    boolRO = True
                    boolVis = False
                Case "CHARSPECIES"
                    str1 = "CHARSPECIES"
                    str2 = "Species"
                    boolRO = False
                    boolVis = True
                Case "CHARSPECIESSTRAIN"
                    str1 = "CHARSPECIESSTRAIN"
                    str2 = "Species Strain"
                    boolRO = False
                    boolVis = True
                Case "CHARDOSEUNITS"
                    str1 = "CHARDOSEUNITS"
                    str2 = "Dose Units"
                    boolRO = False
                    boolVis = True
                Case "CHARDOSECONCUNITS"
                    str1 = "CHARDOSECONCUNITS"
                    str2 = "Dose Conc. Units"
                    boolRO = False
                    boolVis = True
                Case "CHARTISSUEWTUNITS"
                    str1 = "CHARTISSUEWTUNITS"
                    str2 = "Tissue Wt. Units"
                    boolRO = False
                    boolVis = True
                Case "CHARPREVPATREQ"
                    str1 = "CHARPREVPATREQ"
                    str2 = "Previous Rat Requisition"
                    boolRO = False
                    boolVis = False
                Case "CHARSTUDYDESIGNTYPE"
                    str1 = "CHARSTUDYDESIGNTYPE"
                    str2 = "Design Type"
                    boolRO = False
                    boolVis = True
                Case "DTEXTRACTIONDATE"
                    str1 = "DTEXTRACTIONDATE"
                    str2 = "Extraction Date"
                    boolRO = False
                    boolVis = True
                Case "CHARNOTEBOOKREF"
                    str1 = "CHARNOTEBOOKREF"
                    str2 = "Notebook Reference"
                    boolRO = False
                    boolVis = True
                Case "DTASSAYSTARTPRE"
                    str1 = "DTASSAYSTARTPRE"
                    str2 = "Assay Start Predicted"
                    boolRO = False
                    boolVis = True
                Case "DTASSAYSTARTACT"
                    str1 = "DTASSAYSTARTACT"
                    str2 = "Assay Start Actual"
                    boolRO = False
                    boolVis = True
                Case "DTASSAYENDPRED"
                    str1 = "DTASSAYENDPRED"
                    str2 = "Assay End Predicted"
                    boolRO = False
                    boolVis = True
                Case "DTASSAYENDACT"
                    str1 = "DTASSAYENDACT"
                    str2 = "Assay End Actual"
                    boolRO = False
                    boolVis = True
                Case "CHARUSERIDINIT"
                    str1 = "CHARUSERIDINIT"
                    str2 = "User ID Intialized"
                    boolRO = True
                    boolVis = False
                Case "CHARUSERNAMEINIT"
                    str1 = "CHARUSERNAMEINIT"
                    str2 = "User Name Initialized"
                    boolRO = True
                    boolVis = False
                Case "DTINIT"
                    str1 = "DTINIT"
                    str2 = "Date Initialized"
                    boolRO = True
                    boolVis = False
                Case "CHARUSERIDMOD"
                    str1 = "CHARUSERIDMOD"
                    str2 = "User ID Modified"
                    boolRO = True
                    boolVis = False
                Case "CHARUSERNAMEMOD"
                    str1 = "CHARUSERNAMEMOD"
                    str2 = "User Name Modified"
                    boolRO = True
                    boolVis = False
                Case "DTMOD"
                    str1 = "DTMOD"
                    str2 = "Date Modified"
                    boolRO = True
                    boolVis = False
                Case "BOOLINCLUDE"
                    str1 = "BOOLINCLUDE"
                    str2 = "Included"
                    boolRO = True
                    boolVis = False
                Case "BOOLACCEPTED"
                    str1 = "BOOLACCEPTED"
                    str2 = "Accepted"
                    boolRO = True
                    boolVis = False

            End Select
            If Len(str1) = 0 Then
            Else
                'dgv.Columns.Item(str1).ReadOnly = boolRO
                intRow = FindRow(Me.dgvAss, str1)
                If intRow = -1 Then
                Else
                    Me.dgvAss.Rows(intRow).Cells("ColumnValue").ReadOnly = boolRO
                End If
            End If
        Next

end1:

    End Sub

    Sub ConfigSubjectTableSDdv()
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim id As Long
        Dim intRow As Short
        Dim strF As String
        Dim strS As String
        Dim boolGo As Boolean
        Dim int1 As Short
        Dim int2 As Short


        'filter dv for chosen project
        boolGo = True

        dgv = Me.dgvRoutes

        If dgv.RowCount = 0 Then
            boolGo = False
            'GoTo end1
        ElseIf dgv.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = dgv.CurrentRow.Index
        End If
        If boolGo Then
            id = dgv.Item("ID_TBLGUWUPKROUTES", intRow).Value
        Else
            id = -1
        End If
        strF = "ID_TBLGUWUPKROUTES = " & id
        'strS = "NUMTIMEPOINT ASC"

        dv = New DataView(tblGuWuPKSubjects, strF, Nothing, DataViewRowState.CurrentRows)
        dv.AllowDelete = False
        dv.AllowEdit = False
        dv.AllowNew = False

        Me.dgvPatients.DataSource = dv
        int1 = dv.Count 'debugging
        int2 = int1

        Call DoLabel(Me.dgvPatients, Me.lblPatients, "Patients")

        'now fill patients
        Call FillPatient(dv)

end1:

    End Sub

    Sub ConfigSubjectTableSD(ByVal boolW As Boolean)
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        Dim boolRO As Boolean
        'Dim rs As New ADODB.Recordset
        Dim int1 As Short
        Dim intCol As Short
        Dim boolVis As Boolean
        Dim varA

        Call AddPatientColumns()

        Call ConfigSubjectTableSDdv()

        If Me.dgvRoutes.RowCount = 0 Then
            GoTo end1
        End If

        dgv = Me.dgvPatients
        'hide all columns
        For Count1 = 0 To dgv.Columns.Count - 1
            dgv.Columns.Item(Count1).Visible = False
            'dgv.Columns.Item(Count1).DisplayIndex = dgv.Columns.Count - 1
            'dgv.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        'enter column headertexts
        For Count1 = 0 To dgv.ColumnCount - 1
            str4 = dgv.Columns(Count1).Name
            str1 = ""
            boolRO = True
            varA = DataGridViewContentAlignment.BottomLeft
            Select Case str4
                Case "ID_TBLGUWUPKSUBJECTS"
                    str1 = "ID_TBLGUWUPKSUBJECTS"
                    str2 = "ID_TBLGUWUPKSUBJECTS"
                    boolVis = False
                Case "ID_TBLGUWUSTUDIES"
                    str1 = "ID_TBLGUWUSTUDIES"
                    str2 = "ID_TBLGUWUSTUDIES"
                    boolVis = False
                Case "ID_TBLSTUDIES"
                    str1 = "ID_TBLSTUDIES"
                    str2 = "ID_TBLSTUDIES"
                    boolVis = False
                Case "ID_TBLGUWUASSAY"
                    str1 = "ID_TBLGUWUASSAY"
                    str2 = "ID_TBLGUWUASSAY"
                    boolVis = False
                Case "ID_TBLGUWUPKGROUPS"
                    str1 = "ID_TBLGUWUPKGROUPS"
                    str2 = "ID_TBLGUWUPKGROUPS"
                    boolVis = False
                Case "ID_TBLGUWUPKROUTES"
                    str1 = "ID_TBLGUWUPKROUTES"
                    str2 = "ID_TBLGUWUPKROUTES"
                    boolVis = False
                Case "CHARSUBJECTNAME"
                    str1 = "CHARSUBJECTNAME"
                    str2 = "Patient ID"
                    boolVis = True
                    intCol = Count1
                    varA = DataGridViewContentAlignment.BottomLeft
                Case "BOOLTERMINALBLEED"
                    str1 = "BOOLTERMINALBLEED"
                    str2 = "Terminal" & ChrW(10) & "Bleed"
                    boolVis = False
                    varA = DataGridViewContentAlignment.MiddleCenter

                Case "BOOLSERIALBLEED"
                    str1 = "BOOLSERIALBLEED"
                    str2 = "Serial" & ChrW(10) & "Bleed"
                    boolVis = True
                    varA = DataGridViewContentAlignment.MiddleCenter

                Case "NUMPATIENTGROUP"
                    str1 = "NUMPATIENTGROUP"
                    str2 = "Patient" & ChrW(10) & "Group"
                    boolVis = False 'True
                    varA = DataGridViewContentAlignment.MiddleCenter

                Case "NUMTIMEPOINT"
                    str1 = "NUMTIMEPOINT"
                    str2 = "Time" & ChrW(10) & "Point"
                    boolVis = True
                    varA = DataGridViewContentAlignment.MiddleCenter

                Case "CHARUNIQUEID"
                    str1 = "CHARUNIQUEID"
                    str2 = "Unique ID"
                    boolVis = True
                    varA = DataGridViewContentAlignment.BottomLeft

                Case "BOOLTERMINAL"
                    str1 = "BOOLTERMINAL"
                    str2 = "Terminal" & ChrW(10) & "Bleed"
                    boolVis = True
                    varA = DataGridViewContentAlignment.MiddleCenter

                    'Case "BOOLSERIAL"
                    '    str1 = "BOOLSERIAL"
                    '    str2 = "Serial" & ChrW(10) & "Bleed"
                    '    boolVis = True
                    '    varA = DataGridViewContentAlignment.MiddleCenter

            End Select
            If Len(str1) = 0 Then
            Else
                dgv.Columns.Item(str1).Visible = boolVis
                dgv.Columns.Item(str1).HeaderText = str2
                dgv.Columns.Item(str1).ReadOnly = boolRO
                dgv.Columns.Item(str1).DefaultCellStyle.Alignment = varA
                dgv.Columns.Item(str1).SortMode = DataGridViewColumnSortMode.NotSortable

            End If
        Next

        'set first row as current row
        If dgv.Rows.Count = 0 Then

        Else
            dgv.CurrentCell = dgv.Rows.Item(0).Cells(intCol)
        End If
        dgv.AutoResizeColumns()

end1:

    End Sub

    Sub ConfigTimePointTableSDdv()
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim id As Long
        Dim intRow As Short
        Dim strF As String
        Dim strS As String
        Dim boolGo As Boolean
        Dim int1 As Short
        Dim int2 As Short


        'filter dv for chosen project
        boolGo = True

        dgv = Me.dgvRoutes

        If dgv.RowCount = 0 Then
            boolGo = False
            'GoTo end1
        ElseIf dgv.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = dgv.CurrentRow.Index
        End If
        If boolGo Then
            id = dgv.Item("ID_TBLGUWUPKROUTES", intRow).Value
        Else
            id = -1
        End If
        strF = "ID_TBLGUWUPKROUTES = " & id
        strS = "NUMTIMEPOINT ASC"

        dv = New DataView(tblGuWuRTTimePoints, strF, strS, DataViewRowState.CurrentRows)
        dv.AllowDelete = False
        dv.AllowEdit = False
        dv.AllowNew = False

        Me.dgvGroupTimePoints.DataSource = dv
        int1 = dv.Count 'debugging
        int2 = int1

        Call DoLabel(Me.dgvGroupTimePoints, Me.lblTimePoints, "Time Points")


end1:

    End Sub

    Sub ConfigTimePointTableSD(ByVal boolW As Boolean)
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        Dim boolRO As Boolean
        'Dim rs As New ADODB.Recordset
        Dim int1 As Short
        Dim intCol As Short
        Dim boolVis As Boolean

        Call ConfigTimePointTableSDdv()

        If Me.dgvRoutes.RowCount = 0 Then
            GoTo end1
        End If

        dgv = Me.dgvGroupTimePoints
        'hide all columns
        For Count1 = 0 To dgv.Columns.Count - 1
            dgv.Columns.Item(Count1).Visible = False
            'dgv.Columns.Item(Count1).DisplayIndex = dgv.Columns.Count - 1
            'dgv.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        'enter column headertexts
        For Count1 = 0 To dgv.ColumnCount - 1
            str4 = dgv.Columns(Count1).Name
            str1 = ""
            boolRO = False
            Select Case str4
                Case "ID_TBLGUWUPKGROUPS"
                    str1 = "ID_TBLGUWUPKGROUPS"
                    str2 = "ID_TBLGUWUPKGROUPS"
                    boolRO = True
                    boolVis = False
                Case "ID_TBLGUWUSTUDIES"
                    str1 = "ID_TBLGUWUSTUDIES"
                    str2 = "ID_TBLGUWUSTUDIES"
                    boolRO = True
                    boolVis = False
                Case "ID_TBLGUWUASSAY"
                    str1 = "ID_TBLGUWUASSAY"
                    str2 = "ID_TBLGUWUASSAY"
                    boolRO = True
                    boolVis = False
                Case "ID_TBLGUWUPKGROUPS"
                    str1 = "ID_TBLGUWUPKGROUPS"
                    str2 = "ID_TBLGUWUPKGROUPS"
                    boolRO = True
                    boolVis = False
                Case "ID_TBLGUWUPKROUTES"
                    str1 = "ID_TBLGUWUPKROUTES"
                    str2 = "ID_TBLGUWUPKROUTES"
                    boolRO = True
                    boolVis = False

                Case "NUMTIMEPOINT"
                    str1 = "NUMTIMEPOINT"
                    str2 = "Time" & ChrW(10) & "Point" & ChrW(10) & "(hrs)"
                    str2 = "(hrs)"
                    boolRO = True
                    boolVis = True
                    intCol = Count1
            End Select
            If Len(str1) = 0 Then
            Else
                dgv.Columns.Item(str1).Visible = boolVis
                dgv.Columns.Item(str1).HeaderText = str2
                dgv.Columns.Item(str1).ReadOnly = boolRO
            End If
        Next

        'set first row as current row
        If dgv.Rows.Count = 0 Then

        Else
            dgv.CurrentCell = dgv.Rows.Item(0).Cells(intCol)
        End If
        dgv.AutoResizeColumns()

end1:

    End Sub

    Sub ConfigRouteTableSDdv()
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim id As Long
        Dim intRow As Short
        Dim strF As String
        Dim strS As String
        Dim boolGo As Boolean


        'filter dv for chosen project
        boolGo = True
        If Me.dgvGroups.RowCount = 0 Then
            boolGo = False
            'GoTo end1
        ElseIf Me.dgvGroups.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = Me.dgvGroups.CurrentRow.Index
        End If
        If boolGo Then
            id = Me.dgvGroups.Item("ID_TBLGUWUPKGROUPS", intRow).Value
        Else
            id = -1
        End If
        strF = "ID_TBLGUWUPKGROUPS = " & id
        strS = "CHARROUTE ASC"

        dv = New DataView(tblGuWuPKRoutes, strF, strS, DataViewRowState.CurrentRows)
        dv.AllowDelete = False
        dv.AllowEdit = False
        dv.AllowNew = False

        dgv = Me.dgvRoutes

        dgv.DataSource = dv

end1:

    End Sub

    Sub ConfigRouteTableSD(ByVal boolW As Boolean)
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        Dim boolRO As Boolean
        'Dim rs As New ADODB.Recordset
        Dim int1 As Short
        Dim intCol As Short
        Dim boolVis As Boolean
        Dim intRow1 As Short
        Dim var1, var2


        Call ConfigRouteTableSDdv()

        If Me.dgvGroups.RowCount = 0 Then
            GoTo end1
        End If


        dgv = Me.dgvRoutes
        'hide all columns
        For Count1 = 0 To dgv.Columns.Count - 1
            dgv.Columns.Item(Count1).Visible = False
            'dgv.Columns.Item(Count1).DisplayIndex = dgv.Columns.Count - 1
            'dgv.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        'enter column headertexts
        For Count1 = 0 To dgv.ColumnCount - 1
            str4 = dgv.Columns(Count1).Name
            str1 = ""
            boolRO = False
            Select Case str4
                Case "ID_TBLGUWUPKROUTES"
                    str1 = "ID_TBLGUWUPKROUTES"
                    str2 = "ID_TBLGUWUPKROUTES"
                    boolRO = True
                    boolVis = False
                Case "ID_TBLGUWUSTUDIES"
                    str1 = "ID_TBLGUWUSTUDIES"
                    str2 = "ID_TBLGUWUSTUDIES"
                    boolRO = True
                    boolVis = False
                Case "ID_TBLGUWUASSAY"
                    str1 = "ID_TBLGUWUASSAY"
                    str2 = "ID_TBLGUWUASSAY"
                    boolRO = True
                    boolVis = False
                Case "ID_TBLGUWUPKGROUPS"
                    str1 = "ID_TBLGUWUPKGROUPS"
                    str2 = "ID_TBLGUWUPKGROUPS"
                    boolRO = True
                    boolVis = False
                Case "CHARROUTE"
                    str1 = "CHARROUTE"
                    str2 = "Route"
                    boolRO = True
                    boolVis = True
                    intCol = Count1
                Case "CHARTARGETDOSE"
                    str1 = "CHARTARGETDOSE"
                    'find target dose units
                    intRow1 = FindRow(Me.dgvAss, "CHARDOSEUNITS")
                    If intRow1 = -1 Then
                        var1 = ""
                    Else
                        var1 = NZ(Me.dgvAss.Item("ColumnValue", intRow1).Value, "")
                    End If
                    If Len(var1) = 0 Then
                        str2 = "Target Dose"
                    Else
                        str2 = "Target Dose (" & var1 & ")"
                    End If
                    boolRO = False
                    boolVis = False
                Case "CHARTARGETDOSECONC"
                    str1 = "CHARTARGETDOSECONC"
                    'find  dose conc units
                    intRow1 = FindRow(Me.dgvAss, "CHARDOSECONCUNITS")
                    If intRow1 = -1 Then
                        var1 = ""
                    Else
                        var1 = NZ(Me.dgvAss.Item("ColumnValue", intRow1).Value, "")
                    End If
                    If Len(var1) = 0 Then
                        str2 = "Dose Conc."
                    Else
                        str2 = "Dose Conc. (" & var1 & ")"
                    End If
                    boolRO = False
                    boolVis = False
                Case "CHARTARGETTISSUEWT"
                    str1 = "CHARTARGETTISSUEWT"
                    'find  dose conc units
                    intRow1 = FindRow(Me.dgvAss, "CHARTISSUEWTUNITS")
                    If intRow1 = -1 Then
                        var1 = ""
                    Else
                        var1 = NZ(Me.dgvAss.Item("ColumnValue", intRow1).Value, "")
                    End If
                    If Len(var1) = 0 Then
                        str2 = "Target Tissue Wt."
                    Else
                        str2 = "Target Tissue Wt. (" & var1 & ")"
                    End If
                    boolRO = False
                    boolVis = False
                Case "CHARVEHICLE"
                    str1 = "CHARVEHICLE"
                    str2 = "Vehicle"
                    boolRO = False
                    boolVis = False
                Case "CHARFORMULATION"
                    str1 = "CHARFORMULATION"
                    str2 = "Formulation"
                    boolRO = False
                    boolVis = False
                Case "CHARREGIMEN"
                    str1 = "CHARREGIMEN"
                    str2 = "Regimen"
                    boolRO = False
                    boolVis = False
                Case "CHARFASTED"
                    str1 = "CHARFASTED"
                    str2 = "Fasted"
                    boolRO = False
                    boolVis = False
                Case "CHARRESTRAINED"
                    str1 = "CHARRESTRAINED"
                    str2 = "Restrained"
                    boolRO = False
                    boolVis = False
                Case "CHARGENDER"
                    str1 = "CHARGENDER"
                    str2 = "Gender"
                    boolRO = False
                    boolVis = False
                Case "CHARMATRIX"
                    str1 = "CHARMATRIX"
                    str2 = "Collected Matrix"
                    boolRO = False
                    boolVis = False
                Case "CHARTISSUE"
                    str1 = "CHARTISSUE"
                    str2 = "Tissue"
                    boolRO = False
                    boolVis = False
                Case "CHARTARGETTISSUEWT"
                    str1 = "CHARTARGETTISSUEWT"
                    str2 = "Target Tissue Wt"
                    boolRO = False
                    boolVis = False

            End Select
            If Len(str1) = 0 Then
            Else
                dgv.Columns.Item(str1).Visible = boolVis
                dgv.Columns.Item(str1).HeaderText = str2
                dgv.Columns.Item(str1).ReadOnly = boolRO
            End If
        Next

        'set first row as current row
        If dgv.Rows.Count = 0 Then

        Else
            dgv.CurrentCell = dgv.Rows.Item(0).Cells(intCol)
        End If
        dgv.AutoResizeColumns()

end1:

    End Sub

    Sub ConfigGroupTableSDdv()
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim id As Long
        Dim intRow As Short
        Dim strF As String
        Dim strS As String
        Dim boolGo As Boolean


        'filter dv for chosen project
        boolGo = True
        If Me.dgvAssays.RowCount = 0 Then
            boolGo = False
            'GoTo end1
        ElseIf Me.dgvAssays.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = Me.dgvAssays.CurrentRow.Index
        End If
        If boolGo Then
            id = Me.dgvAssays.Item("ID_TBLGUWUASSAY", intRow).Value
        Else
            id = -1
        End If
        strF = "ID_TBLGUWUASSAY = " & id
        strS = "CHARGROUP ASC"

        dv = New DataView(tblGuWuPKGroups, strF, strS, DataViewRowState.CurrentRows)
        dv.AllowDelete = False
        dv.AllowEdit = False
        dv.AllowNew = False

        dgv = Me.dgvGroups

        dgv.DataSource = dv

end1:

    End Sub

    Sub ConfigCmpdLotTableSDdv()
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim id As Long
        Dim intRow As Short
        Dim strF As String
        Dim strS As String
        Dim boolGo As Boolean


        'filter dv for chosen project
        boolGo = True
        If Me.dgvAssays.RowCount = 0 Then
            boolGo = False
            'GoTo end1
        ElseIf Me.dgvAssays.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = Me.dgvAssays.CurrentRow.Index
        End If
        If boolGo Then
            id = Me.dgvAssays.Item("ID_TBLGUWUASSAY", intRow).Value
        Else
            id = -1
        End If
        strF = "ID_TBLGUWUASSAY = " & id
        'strS = "CHARCOMPOUND ASC"

        Try
            dv = New DataView(tblGuWuAssignedCmpdLot, strF, Nothing, DataViewRowState.CurrentRows)
            dv.AllowDelete = False
            dv.AllowEdit = False
            dv.AllowNew = False

            dgv = Me.dgvLotNum

            dgv.DataSource = dv

            'now fill LOTS
            Call FillCmpdLot(dv)


        Catch ex As Exception

        End Try

end1:

    End Sub

    Sub ConfigPersTableSDdv()
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim dv1 As system.data.dataview
        Dim id As Long
        Dim intRow As Short
        Dim strF As String
        Dim strS As String
        Dim boolGo As Boolean


        'filter dv for chosen project
        boolGo = True
        If Me.dgvAssays.RowCount = 0 Then
            boolGo = False
            'GoTo end1
        ElseIf Me.dgvAssays.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = Me.dgvAssays.CurrentRow.Index
        End If
        If boolGo Then
            id = Me.dgvAssays.Item("ID_TBLGUWUASSAY", intRow).Value
        Else
            id = -1
        End If

        'first do PI

        strF = "ID_TBLGUWUASSAY = " & id & " AND CHARROLE = 'PI' AND DTREMOVED IS NULL"
        strS = "CHARPERSONNEL ASC"

        dv = New DataView(tblGuWuAssayPERS, strF, strS, DataViewRowState.CurrentRows)
        dv.AllowDelete = False
        dv.AllowEdit = False
        dv.AllowNew = False

        dgv = Me.dgvPI

        dgv.DataSource = dv

        'now do Analyst
        strF = "ID_TBLGUWUASSAY = " & id & " AND CHARROLE = 'Analyst' AND DTREMOVED IS NULL"
        strS = "CHARPERSONNEL ASC"

        dv = New DataView(tblGuWuAssayPERS, strF, strS, DataViewRowState.CurrentRows)
        dv.AllowDelete = False
        dv.AllowEdit = False
        dv.AllowNew = False

        dgv = Me.dgvAnalyst

        dgv.DataSource = dv

end1:

    End Sub

    Sub ConfigPITableSDdv()
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim id As Long
        Dim intRow As Short
        Dim strF As String
        Dim strS As String
        Dim boolGo As Boolean


        'filter dv for chosen project
        boolGo = True
        If Me.dgvAssays.RowCount = 0 Then
            boolGo = False
            'GoTo end1
        ElseIf Me.dgvAssays.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = Me.dgvAssays.CurrentRow.Index
        End If
        If boolGo Then
            id = Me.dgvAssays.Item("ID_TBLGUWUASSAY", intRow).Value
        Else
            id = -1
        End If
        strF = "ID_TBLGUWUASSAY = " & id & " AND CHARROLE = 'PI' AND DTREMOVED IS NULL"
        strS = "CHARPERSONNEL ASC"

        dv = New DataView(tblGuWuAssayPERS, strF, strS, DataViewRowState.CurrentRows)
        dv.AllowDelete = False
        dv.AllowEdit = False
        dv.AllowNew = False

        dgv = Me.dgvPI

        dgv.DataSource = dv

        'now fill personnel
        Call FillPersonnel(dv)


end1:

    End Sub

    Sub ConfigAnalystTableSDdv()

        Dim dgv As DataGridView
        Dim dv As System.Data.DataView
        Dim id As Long
        Dim intRow As Short
        Dim strF As String
        Dim strS As String
        Dim boolGo As Boolean


        'filter dv for chosen project
        boolGo = True
        If Me.dgvAssays.RowCount = 0 Then
            boolGo = False
            'GoTo end1
        ElseIf Me.dgvAssays.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = Me.dgvAssays.CurrentRow.Index
        End If
        If boolGo Then
            id = Me.dgvAssays.Item("ID_TBLGUWUASSAY", intRow).Value
        Else
            id = -1
        End If
        strF = "ID_TBLGUWUASSAY = " & id & " AND CHARROLE = 'Analyst' AND DTREMOVED IS NULL"
        strS = "CHARPERSONNEL ASC"

        dv = New DataView(tblGuWuAssayPERS, strF, strS, DataViewRowState.CurrentRows)
        dv.AllowDelete = False
        dv.AllowEdit = False
        dv.AllowNew = False

        dgv = Me.dgvAnalyst

        dgv.DataSource = dv

        'now fill personnel
        Call FillPersonnel(dv)


end1:

    End Sub

    Sub FillPersonnel(ByVal dv As System.Data.DataView)

        Dim dv1 As system.data.dataview = New DataView(tblPersonnel)
        Dim strF As String
        Dim strS As String
        Dim intRows As Int16
        Dim Count1 As Int16
        Dim id As Int64
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim strName As String
        Dim boolEdit As Boolean

        boolEdit = dv.AllowEdit
        dv.AllowEdit = True

        Try

        Catch ex As Exception

        End Try

        intRows = dv.Count
        For Count1 = 0 To intRows - 1
            id = dv(Count1).Item("ID_TBLPERSONNEL")
            strF = "ID_TBLPERSONNEL = " & id
            dv1.RowFilter = strF

            str1 = NZ(dv1(0).Item("CHARFIRSTNAME"), "")
            str2 = NZ(dv1(0).Item("CHARMIDDLENAME"), "")
            str3 = NZ(dv1(0).Item("CHARLASTNAME"), "")

            If Len(str2) = 0 Then
                strName = str3 & ", " & str1
            Else
                strName = str3 & ", " & str1 & " " & str2
            End If

            dv(Count1).BeginEdit()
            dv(Count1).Item("CHARPERSONNEL") = strName
            dv(Count1).EndEdit()

        Next

        dv.AllowEdit = boolEdit

        tblGuWuAssayPERS.AcceptChanges()


    End Sub

    Sub AddPersColumns()

        'add an unbound column to tblguwuassignedcmpd

        Dim dtbl As System.Data.Datatable

        dtbl = tblGuWuAssayPERS

        If dtbl.Columns.Contains("CHARPERSONNEL") Then
            Exit Sub
        End If

        Dim col1 As New DataColumn
        col1.AllowDBNull = True
        col1.ColumnName = "CHARPERSONNEL"
        col1.DataType = System.Type.GetType("System.String")
        col1.Caption = "Name"
        dtbl.Columns.Add(col1)

    End Sub

    Sub ConfigPersTableSD(ByVal boolW As Boolean)
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        Dim boolRO As Boolean
        'Dim rs As New ADODB.Recordset
        Dim int1 As Short
        Dim intCol As Short
        Dim boolVis As Boolean

        Call AddPersColumns()

        Call ConfigPITableSDdv()
        Call ConfigAnalystTableSDdv()

        If Me.dgvAssays.RowCount = 0 Then
            GoTo end1
        End If

        'first do PI
        dgv = Me.dgvPI
        'hide all columns
        For Count1 = 0 To dgv.Columns.Count - 1
            dgv.Columns.Item(Count1).Visible = False
            dgv.Columns.Item(Count1).ReadOnly = True
            'dgv.Columns.Item(Count1).DisplayIndex = dgv.Columns.Count - 1
            'dgv.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        'enter column headertexts
        For Count1 = 0 To dgv.ColumnCount - 1
            str4 = dgv.Columns(Count1).Name
            str1 = ""
            boolRO = False
            Select Case str4
                Case "CHARPERSONNEL"
                    str1 = "CHARPERSONNEL"
                    str2 = "Name"
                    boolRO = True
                    boolVis = True
                    intCol = Count1
            End Select
            If Len(str1) = 0 Then
            Else
                dgv.Columns.Item(str1).Visible = boolVis
                dgv.Columns.Item(str1).HeaderText = str2
                dgv.Columns.Item(str1).ReadOnly = boolRO
            End If
        Next

        'set first row as current row
        If dgv.Rows.Count = 0 Then

        Else
            dgv.CurrentCell = dgv.Rows.Item(0).Cells(intCol)
        End If
        dgv.AutoResizeColumns()


        'now do Analyst
        dgv = Me.dgvAnalyst
        'hide all columns
        For Count1 = 0 To dgv.Columns.Count - 1
            dgv.Columns.Item(Count1).Visible = False
            dgv.Columns.Item(Count1).ReadOnly = True
            'dgv.Columns.Item(Count1).DisplayIndex = dgv.Columns.Count - 1
            'dgv.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        'enter column headertexts
        For Count1 = 0 To dgv.ColumnCount - 1
            str4 = dgv.Columns(Count1).Name
            str1 = ""
            boolRO = False
            Select Case str4
                Case "CHARPERSONNEL"
                    str1 = "CHARPERSONNEL"
                    str2 = "Name"
                    boolRO = True
                    boolVis = True
                    intCol = Count1
            End Select
            If Len(str1) = 0 Then
            Else
                dgv.Columns.Item(str1).Visible = boolVis
                dgv.Columns.Item(str1).HeaderText = str2
                dgv.Columns.Item(str1).ReadOnly = boolRO
            End If
        Next

        'set first row as current row
        If dgv.Rows.Count = 0 Then

        Else
            dgv.CurrentCell = dgv.Rows.Item(0).Cells(intCol)
        End If
        dgv.AutoResizeColumns()

end1:

    End Sub

    Sub ConfigCmpdTableSDdv()
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim id As Long
        Dim intRow As Short
        Dim strF As String
        Dim strS As String
        Dim boolGo As Boolean


        'filter dv for chosen project
        boolGo = True
        If Me.dgvAssays.RowCount = 0 Then
            boolGo = False
            'GoTo end1
        ElseIf Me.dgvAssays.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = Me.dgvAssays.CurrentRow.Index
        End If
        If boolGo Then
            id = Me.dgvAssays.Item("ID_TBLGUWUASSAY", intRow).Value
        Else
            id = -1
        End If
        strF = "ID_TBLGUWUASSAY = " & id
        strS = "CHARCOMPOUND ASC"

        dv = New DataView(tblGuWuAssignedCmpd, strF, strS, DataViewRowState.CurrentRows)
        dv.AllowDelete = False
        dv.AllowEdit = False
        dv.AllowNew = False

        dgv = Me.dgvCmpd

        dgv.DataSource = dv

        'now fill compounds
        Call FillCmpd(dv)


end1:

    End Sub

    Sub FillCmpdLot(ByVal dv As System.Data.DataView)

        Dim tbl As System.Data.Datatable
        Dim tbl1 As System.Data.Datatable
        Dim rows() As DataRow
        Dim rows1() As DataRow

        Dim strF As String
        Dim strF1 As String
        Dim strS As String
        Dim intRows As Int16
        Dim Count1 As Int16
        Dim id As Int64
        Dim id1 As Int64
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim strName As String
        Dim arr1(3, 1) As Int64

        intRows = dv.Count
        ReDim arr1(3, intRows)

        tbl = tblGuWuCompoundsInd
        tbl1 = tblGuWuAssignedCmpdLot

        Dim var1

        For Count1 = 0 To intRows - 1
            'who knows why I have to do this, but I do
            arr1(0, Count1) = dv(Count1).Item("ID_TBLGUWUASSIGNEDCMPDLOT")
            arr1(1, Count1) = dv(Count1).Item("ID_TBLGUWUCOMPOUNDSIND")
        Next

        intRows = dv.Count
        For Count1 = 0 To intRows - 1

            id = arr1(0, Count1)
            id1 = arr1(1, Count1)
            strF = "ID_TBLGUWUCOMPOUNDSIND = " & id1
            Erase rows
            rows = tbl.Select(strF)

            str1 = NZ(rows(0).Item("CHARLOTNUMBER"), "")

            strF1 = "ID_TBLGUWUASSIGNEDCMPDLOT = " & id
            Erase rows
            rows1 = tbl1.Select(strF1)

            rows1(0).BeginEdit()
            rows1(0).Item("CHARLOTNUMBER") = str1
            rows1(0).EndEdit()

            'id = dv(Count1).Item("ID_TBLGUWUCOMPOUNDSIND")
            'strF = "ID_TBLGUWUCOMPOUNDSIND = " & id
            'dv1.RowFilter = strF

            'str1 = NZ(dv1(0).Item("CHARLOTNUMBER"), "")

            'dv(Count1).BeginEdit()
            'dv(Count1).Item("CHARLOTNUMBER") = str1
            'dv(Count1).EndEdit()

        Next

        Me.dgvLotNum.AutoResizeColumns()


    End Sub

    Sub FillPatient(ByVal dv As System.Data.DataView)

        Dim tbl As System.Data.Datatable
        Dim tbl1 As System.Data.Datatable
        Dim rows() As DataRow
        Dim rows1() As DataRow
        Dim strF As String
        Dim strF1 As String
        Dim strS As String
        Dim intRows As Int16
        Dim Count1 As Int16
        Dim id As Int64
        Dim id1 As Int64
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim strName As String
        Dim arr1(3, 1) As Int64
        Dim num1 As Single

        intRows = dv.Count
        ReDim arr1(3, intRows)

        'tbl = tblGuWuCompounds
        'tbl1 = tblGuWuAssignedCmpd

        tbl = tblGuWuPKSubjects
        tbl1 = tblGuWuRTTimePoints

        Dim var1

        For Count1 = 0 To intRows - 1
            'who knows why I have to do this, but I do
            'arr1(0, Count1) = dv(Count1).Item("ID_TBLGUWUASSIGNEDCMPD")
            'arr1(1, Count1) = dv(Count1).Item("ID_TBLGUWUCOMPOUNDS")

            arr1(0, Count1) = NZ(dv(Count1).Item("ID_tblGuWuRTTimePoints"), -1)
            arr1(1, Count1) = dv(Count1).Item("ID_tblGuWuPKSubjects")
        Next

        For Count1 = 0 To intRows - 1
            id = arr1(0, Count1) 'dv(Count1).Item("ID_tblGuWuRTTimePoints")
            id1 = arr1(1, Count1) 'dv(Count1).Item("ID_tblGuWuPKSubjects")
            strF = "ID_tblGuWuRTTimePoints = " & id
            Erase rows1
            rows1 = tbl1.Select(strF)

            If rows1.Length = 0 Then
                num1 = -1
            Else
                num1 = rows1(0).Item("NUMTIMEPOINT")
            End If


            strF1 = "ID_tblGuWuPKSubjects = " & id1
            Erase rows
            rows = tbl.Select(strF1)

            rows(0).BeginEdit()
            rows(0).Item("NUMTIMEPOINT") = num1
            rows(0).EndEdit()

        Next

        Me.dgvPatients.AutoResizeColumns()


    End Sub

    Sub FillCmpd(ByVal dv As System.Data.DataView)

        Dim tbl As System.Data.Datatable
        Dim tbl1 As System.Data.Datatable
        Dim rows() As DataRow
        Dim rows1() As DataRow
        Dim strF As String
        Dim strF1 As String
        Dim strS As String
        Dim intRows As Int16
        Dim Count1 As Int16
        Dim id As Int64
        Dim id1 As Int64
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim strName As String
        Dim arr1(3, 1) As Int64

        intRows = dv.Count
        ReDim arr1(3, intRows)

        tbl = tblGuWuCompounds
        tbl1 = tblGuWuAssignedCmpd

        Dim var1

        For Count1 = 0 To intRows - 1
            'who knows why I have to do this, but I do
            arr1(0, Count1) = dv(Count1).Item("ID_TBLGUWUASSIGNEDCMPD")
            arr1(1, Count1) = dv(Count1).Item("ID_TBLGUWUCOMPOUNDS")
        Next

        For Count1 = 0 To intRows - 1
            id = arr1(0, Count1) 'dv(Count1).Item("ID_TBLGUWUCOMPOUNDS")
            id1 = arr1(1, Count1) 'dv(Count1).Item("ID_TBLGUWUCOMPOUNDS")
            strF = "ID_TBLGUWUCOMPOUNDS = " & id1
            Erase rows
            rows = tbl.Select(strF)

            str1 = NZ(rows(0).Item("CHARANALYTENAME"), "")
            str2 = NZ(rows(0).Item("CHARCOMPANYID"), "")

            strF1 = "ID_TBLGUWUASSIGNEDCMPD = " & id
            Erase rows
            rows1 = tbl1.Select(strF1)

            rows1(0).BeginEdit()
            rows1(0).Item("CHARCOMPOUND") = str1
            rows1(0).Item("CHARCOMPANYID") = str2
            rows1(0).EndEdit()

        Next

        Me.dgvCmpd.AutoResizeColumns()


    End Sub


    Sub AddCmpdLotColumns()

        'add an unbound column to tblguwuassignedcmpd

        Dim dtbl As System.Data.Datatable

        dtbl = tblGuWuAssignedCmpdLot

        If dtbl.Columns.Contains("CHARLOTNUMBER") Then
            Exit Sub
        End If

        Dim col1 As New DataColumn
        col1.AllowDBNull = True
        col1.ColumnName = "CHARLOTNUMBER"
        col1.DataType = System.Type.GetType("System.String")
        col1.Caption = "Lot Number"
        dtbl.Columns.Add(col1)


    End Sub

    Sub AddPatientColumns()

        'add an unbound column to tblguwuassignedcmpd

        Dim dtbl As System.Data.Datatable

        dtbl = tblGuWuPKSubjects

        If dtbl.Columns.Contains("NUMTIMEPOINT") Then
            Exit Sub
        End If

        Dim col1 As New DataColumn
        col1.AllowDBNull = True
        col1.ColumnName = "NUMTIMEPOINT"
        col1.DataType = System.Type.GetType("System.Single")
        col1.Caption = "Time Point (hrs)"
        dtbl.Columns.Add(col1)

        'Dim col2 As New DataColumn
        'col2.AllowDBNull = True
        'col2.ColumnName = "BOOLSERIAL"
        'col2.DataType = System.Type.GetType("System.Boolean")
        'col2.Caption = "Serial" & ChrW(10) & "Bleed"
        'dtbl.Columns.Add(col2)

        Dim col3 As New DataColumn
        col3.AllowDBNull = True
        col3.ColumnName = "BOOLTERMINAL"
        col3.DataType = System.Type.GetType("System.Boolean")
        col3.Caption = "Terminal" & ChrW(10) & "Bleed"
        dtbl.Columns.Add(col3)

        'If Me.dgvPatients.Columns.Contains("BOOLSERIAL") Then
        'Else
        '    'make some columns checkboxes
        '    Dim column As New DataGridViewCheckBoxColumn()
        '    With column
        '        .HeaderText = "Serial" & ChrW(10) & "Bleed"
        '        .Name = "BOOLSERIAL"
        '        .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        '        .FlatStyle = FlatStyle.Standard
        '        .CellTemplate = New DataGridViewCheckBoxCell()
        '        '.CellTemplate.Style.BackColor = Color.Beige
        '    End With
        '    dgv.Columns.Insert(0, column)


        '    Dim column1 As New DataGridViewCheckBoxColumn()
        '    With column1
        '        .HeaderText = "Terminal" & ChrW(10) & "Bleed"
        '        .Name = "BOOLTERMINAL"
        '        .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        '        .FlatStyle = FlatStyle.Standard
        '        .CellTemplate = New DataGridViewCheckBoxCell()
        '        '.CellTemplate.Style.BackColor = Color.Beige
        '    End With
        '    dgv.Columns.Insert(0, column1)
        'End If

    End Sub

    Sub AddCmpdColumns()

        'add an unbound column to tblguwuassignedcmpd

        Dim dtbl As System.Data.Datatable

        dtbl = tblGuWuAssignedCmpd

        If dtbl.Columns.Contains("CHARCOMPOUND") Then
            Exit Sub
        End If

        Dim col1 As New DataColumn
        col1.AllowDBNull = True
        col1.ColumnName = "CHARCOMPOUND"
        col1.DataType = System.Type.GetType("System.String")
        col1.Caption = "Compound"
        dtbl.Columns.Add(col1)

        Dim col2 As New DataColumn
        col2.AllowDBNull = True
        col2.ColumnName = "CHARCOMPANYID"
        col2.DataType = System.Type.GetType("System.String")
        col2.Caption = "Compound ID"
        dtbl.Columns.Add(col2)


    End Sub

    Sub ConfigCmpdTableSD(ByVal boolW As Boolean)

        Dim dgv As DataGridView
        Dim dv As System.Data.DataView
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        Dim boolRO As Boolean
        'Dim rs As New ADODB.Recordset
        Dim int1 As Short
        Dim intCol As Short
        Dim boolVis As Boolean
        Dim intDisplay As Short

        Call AddCmpdColumns()
        Call AddCmpdLotColumns()

        Call Configure_dgvCmpdLot()

        Call ConfigCmpdTableSDdv()
        Call ConfigCmpdLotTableSDdv()

        If Me.dgvAssays.RowCount = 0 Then
            GoTo end1
        End If

        dgv = Me.dgvCmpd
        'hide all columns
        For Count1 = 0 To dgv.Columns.Count - 1
            dgv.Columns.Item(Count1).Visible = False
            dgv.Columns.Item(Count1).ReadOnly = True
            'dgv.Columns.Item(Count1).DisplayIndex = dgv.Columns.Count - 1
            'dgv.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        'enter column headertexts
        For Count1 = 0 To dgv.ColumnCount - 1
            str4 = dgv.Columns(Count1).Name
            str1 = ""
            boolRO = False
            intDisplay = -1
            Select Case str4
                Case "CHARCOMPOUND"
                    str1 = "CHARCOMPOUND"
                    str2 = "Compound"
                    boolRO = True
                    boolVis = True
                    intCol = Count1
                    intDisplay = 0
                Case "CHARCOMPANYID"
                    str1 = "CHARCOMPANYID"
                    str2 = "Company ID"
                    boolRO = True
                    boolVis = True
                    intDisplay = 1
                Case "CHARCOMPOUNDTYPE"
                    str1 = "CHARCOMPOUNDTYPE"
                    str2 = "Type"
                    boolRO = True
                    boolVis = True
                    intDisplay = 2
            End Select
            If Len(str1) = 0 Then
            Else
                dgv.Columns.Item(str1).Visible = boolVis
                dgv.Columns.Item(str1).HeaderText = str2
                dgv.Columns.Item(str1).ReadOnly = boolRO
            End If
            If intDisplay = -1 Then
            Else
                dgv.Columns.Item(str1).DisplayIndex = intDisplay
            End If
        Next

        'set first row as current row
        If dgv.Rows.Count = 0 Then

        Else
            dgv.CurrentCell = dgv.Rows.Item(0).Cells(intCol)
        End If
        dgv.AutoResizeColumns()


        dgv = Me.dgvLotNum
        'hide all columns
        For Count1 = 0 To dgv.Columns.Count - 1
            dgv.Columns.Item(Count1).Visible = False
            dgv.Columns.Item(Count1).ReadOnly = True
            'dgv.Columns.Item(Count1).DisplayIndex = dgv.Columns.Count - 1
            'dgv.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        'enter column headertexts
        For Count1 = 0 To dgv.ColumnCount - 1
            str4 = dgv.Columns(Count1).Name
            str1 = ""
            boolRO = False
            Select Case str4
                Case "CHARLOTNUMBER"
                    str1 = "CHARLOTNUMBER"
                    str2 = "Lot Number"
                    boolRO = True
                    boolVis = True
                    intCol = Count1
            End Select
            If Len(str1) = 0 Then
            Else
                dgv.Columns.Item(str1).Visible = boolVis
                dgv.Columns.Item(str1).HeaderText = str2
                dgv.Columns.Item(str1).ReadOnly = boolRO
            End If
        Next

        dgv.Columns(intCol).MinimumWidth = (dgv.Width - dgv.RowHeadersWidth) * 0.85

        'set first row as current row
        If dgv.Rows.Count = 0 Then

        Else
            dgv.CurrentCell = dgv.Rows.Item(0).Cells(intCol)
        End If
        dgv.AutoResizeColumns()

end1:

    End Sub

    Sub ConfigGroupTableSD(ByVal boolW As Boolean)
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        Dim boolRO As Boolean
        'Dim rs As New ADODB.Recordset
        Dim int1 As Short
        Dim intCol As Short
        Dim boolVis As Boolean

        Call ConfigGroupTableSDdv()

        If Me.dgvAssays.RowCount = 0 Then
            GoTo end1
        End If

        dgv = Me.dgvGroups
        'hide all columns
        For Count1 = 0 To dgv.Columns.Count - 1
            dgv.Columns.Item(Count1).Visible = False
            'dgv.Columns.Item(Count1).DisplayIndex = dgv.Columns.Count - 1
            'dgv.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        'enter column headertexts
        For Count1 = 0 To dgv.ColumnCount - 1
            str4 = dgv.Columns(Count1).Name
            str1 = ""
            boolRO = False
            Select Case str4
                Case "ID_TBLGUWUPKGROUPS"
                    str1 = "ID_TBLGUWUPKGROUPS"
                    str2 = "ID_TBLGUWUPKGROUPS"
                    boolRO = True
                    boolVis = False
                Case "ID_TBLGUWUSTUDIES"
                    str1 = "ID_TBLGUWUSTUDIES"
                    str2 = "ID_TBLGUWUSTUDIES"
                    boolRO = True
                    boolVis = False
                Case "ID_TBLGUWUASSAY"
                    str1 = "ID_TBLGUWUASSAY"
                    str2 = "ID_TBLGUWUASSAY"
                    boolRO = True
                    boolVis = False
                Case "CHARGROUP"
                    str1 = "CHARGROUP"
                    str2 = "Group"
                    boolRO = True
                    boolVis = True
                    intCol = Count1
            End Select
            If Len(str1) = 0 Then
            Else
                dgv.Columns.Item(str1).Visible = boolVis
                dgv.Columns.Item(str1).HeaderText = str2
                dgv.Columns.Item(str1).ReadOnly = boolRO
            End If
        Next

        'set first row as current row
        If dgv.Rows.Count = 0 Then

        Else
            dgv.CurrentCell = dgv.Rows.Item(0).Cells(intCol)
        End If
        dgv.AutoResizeColumns()

end1:

    End Sub

    Sub ConfigAssayTableSDdv()
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim id As Long
        Dim intRow As Short
        Dim strF As String
        Dim strS As String


        'filter dv for chosen project
        If Me.dgvSDStudy.RowCount = 0 Then
            GoTo end1
        ElseIf Me.dgvSDStudy.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = Me.dgvSDStudy.CurrentRow.Index
        End If
        id = Me.dgvSDStudy.Item("ID_TBLGUWUSTUDIES", intRow).Value
        strF = "ID_TBLGUWUSTUDIES = " & id
        strS = "CHARASSAYNAME ASC"

        dv = New DataView(tblGuWuAssay, strF, strS, DataViewRowState.CurrentRows)
        dv.AllowDelete = False
        dv.AllowEdit = False
        dv.AllowNew = False

        dgv = Me.dgvAssays

        dgv.DataSource = dv

end1:

    End Sub

    Sub ConfigAssayTableSD(ByVal boolW As Boolean)
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        Dim boolRO As Boolean
        'Dim rs As New ADODB.Recordset
        Dim int1 As Short
        Dim intCol As Short
        Dim boolVis As Boolean

        If Me.dgvSDStudy.RowCount = 0 Then
            GoTo end1
        End If

        Call ConfigAssayTableSDdv()

        dgv = Me.dgvAssays
        'hide all columns
        For Count1 = 0 To dgv.Columns.Count - 1
            dgv.Columns.Item(Count1).Visible = False
            'dgv.Columns.Item(Count1).DisplayIndex = dgv.Columns.Count - 1
            'dgv.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        'enter column headertexts
        For Count1 = 0 To dgv.ColumnCount - 1
            str4 = dgv.Columns(Count1).Name
            str1 = ""
            boolRO = False
            Select Case str4
                Case "ID_TBLGUWUASSAY"
                    str1 = "ID_TBLGUWUASSAY"
                    str2 = "ID_TBLGUWUASSAY"
                    boolRO = True
                    boolVis = False
                Case "ID_TBLGUWUSTUDIES"
                    str1 = "ID_TBLGUWUSTUDIES"
                    str2 = "ID_TBLGUWUSTUDIES"
                    boolRO = True
                    boolVis = False
                Case "ID_TBLSTUDIES"
                    str1 = "ID_TBLSTUDIES"
                    str2 = "ID_TBLSTUDIES"
                    boolRO = True
                    boolVis = False
                Case "CHARASSAYNAME"
                    str1 = "CHARASSAYNAME"
                    str2 = "Assay Name"
                    boolRO = False
                    boolVis = True
                    intCol = Count1
                Case "DTASSAYDATE"
                    str1 = "DTASSAYDATE"
                    str2 = "Assay Date"
                    boolRO = False
                    boolVis = True
                Case "ID_TBLGUWUSPECIES"
                    str1 = "ID_TBLGUWUSPECIES"
                    str2 = "ID_TBLGUWUSPECIES"
                    boolRO = True
                    boolVis = False
                Case "CHARSPECIES"
                    str1 = "CHARSPECIES"
                    str2 = "Species"
                    boolRO = False
                    boolVis = True
                Case "CHARSPECIESSTRAIN"
                    str1 = "CHARSPECIESSTRAIN"
                    str2 = "Species Strain"
                    boolRO = False
                    boolVis = True
                Case "CHARDOSEUNITS"
                    str1 = "CHARDOSEUNITS"
                    str2 = "Dose Units"
                    boolRO = False
                    boolVis = True
                Case "CHARDOSECONCUNITS"
                    str1 = "CHARDOSECONCUNITS"
                    str2 = "Dose Conc. Units"
                    boolRO = False
                    boolVis = True
                Case "CHARTISSUEWTUNITS"
                    str1 = "CHARTISSUEWTUNITS"
                    str2 = "Tissue Wt. Units"
                    boolRO = False
                    boolVis = True
                Case "CHARPREVPATREQ"
                    str1 = "CHARPREVPATREQ"
                    str2 = "Previous Rat Requisition"
                    boolRO = False
                    boolVis = False
                Case "CHARSTUDYDESIGNTYPE"
                    str1 = "CHARSTUDYDESIGNTYPE"
                    str2 = "Design Type"
                    boolRO = False
                    boolVis = True
                Case "DTEXTRACTIONDATE"
                    str1 = "DTEXTRACTIONDATE"
                    str2 = "Extraction Date"
                    boolRO = False
                    boolVis = True
                Case "CHARNOTEBOOKREF"
                    str1 = "CHARNOTEBOOKREF"
                    str2 = "Notebook Reference"
                    boolRO = False
                    boolVis = True
                Case "DTASSAYSTARTPRE"
                    str1 = "DTASSAYSTARTPRE"
                    str2 = "Assay Start Predicted"
                    boolRO = False
                    boolVis = True
                Case "DTASSAYSTARTACT"
                    str1 = "DTASSAYSTARTACT"
                    str2 = "Assay Start Actual"
                    boolRO = False
                    boolVis = True
                Case "DTASSAYENDPRED"
                    str1 = "DTASSAYENDPRED"
                    str2 = "Assay End Predicted"
                    boolRO = False
                    boolVis = True
                Case "DTASSAYENDACT"
                    str1 = "DTASSAYENDACT"
                    str2 = "Assay End Actual"
                    boolRO = False
                    boolVis = True
                Case "CHARUSERIDINIT"
                    str1 = "CHARUSERIDINIT"
                    str2 = "User ID Intialized"
                    boolRO = True
                    boolVis = False
                Case "CHARUSERNAMEINIT"
                    str1 = "CHARUSERNAMEINIT"
                    str2 = "User Name Initialized"
                    boolRO = True
                    boolVis = False
                Case "DTINIT"
                    str1 = "DTINIT"
                    str2 = "Date Initialized"
                    boolRO = True
                    boolVis = False
                Case "CHARUSERIDMOD"
                    str1 = "CHARUSERIDMOD"
                    str2 = "User ID Modified"
                    boolRO = True
                    boolVis = False
                Case "CHARUSERNAMEMOD"
                    str1 = "CHARUSERNAMEMOD"
                    str2 = "User Name Modified"
                    boolRO = True
                    boolVis = False
                Case "DTMOD"
                    str1 = "DTMOD"
                    str2 = "Date Modified"
                    boolRO = True
                    boolVis = False
                Case "BOOLINCLUDE"
                    str1 = "BOOLINCLUDE"
                    str2 = "Included"
                    boolRO = True
                    boolVis = False
                Case "BOOLACCEPTED"
                    str1 = "BOOLACCEPTED"
                    str2 = "Accepted"
                    boolRO = True
                    boolVis = False


            End Select
            If Len(str1) = 0 Then
            Else
                dgv.Columns.Item(str1).Visible = boolVis
                dgv.Columns.Item(str1).HeaderText = str2
                dgv.Columns.Item(str1).ReadOnly = boolRO
            End If
        Next

        'set first row as current row
        If dgv.Rows.Count = 0 Then

        Else
            dgv.CurrentCell = dgv.Rows.Item(0).Cells(intCol)
        End If
        dgv.AutoResizeColumns()

end1:

    End Sub

    Private Sub rbGuWu_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbGuWu.CheckedChanged
        If boolFormLoad Then
            Exit Sub
        End If

        If Me.rbGuWu.Checked Then
            Me.gbxProjectView.Visible = True
            Me.gbxStudyView.Visible = True
            Me.panCP.Visible = True
        Else
            Me.gbxProjectView.Visible = False
            Me.gbxStudyView.Visible = False
            Me.panCP.Visible = False
        End If

        Call SDDataSourceChecked()

        Call ShowGuWudgv()

        Call ChangeStudies()

        If Me.rbGuWu.Checked Then
            Me.cmdEdit.Enabled = boolCmdEditE
        Else
            Me.cmdEdit.Enabled = False
        End If


    End Sub

    Private Sub dgvTab1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvTab1.SelectionChanged
        If boolHold Then
            Exit Sub
        End If
        boolHold = True
        Call SelectTab1(1)
        boolHold = False

        'pesky
        Call DoLabel(Me.dgvAssays, Me.lbldgvAssays, "Assays")

    End Sub

    Private Sub sst1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles sst1.Click
        If boolHold Then
            Exit Sub
        End If
        boolHold = True

        Call SelectTab1(2)

        boolHold = False
    End Sub

    Sub SaveSDProjects()

        Me.dgvProj.CommitEdit(DataGridViewDataErrorContexts.Commit)

        Call UpdateDGVs("Projects")

        Call FillUserNames("Projects")

        If boolGuWuOracle Then
            'Try
            '    ta_TBLGUWUPROJECTS.Update(tblGuWuProjects)
            'Catch ex As DBConcurrencyException
            '    'ds2005.tblGuWuProjects.Merge('ds2005.tblGuWuProjects, True)
            'End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_TBLGUWUPROJECTSAcc.Update(tblGuWuProjects)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLGUWUPROJECTS.Merge('ds2005Acc.TBLGUWUPROJECTS, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_TBLGUWUPROJECTSSQLServer.Update(tblGuWuProjects)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLGUWUPROJECTS.Merge('ds2005Acc.TBLGUWUPROJECTS, True)
            End Try
        End If

    End Sub

    Sub SaveSDStudies()

        Me.dgvStud.CommitEdit(DataGridViewDataErrorContexts.Commit)

        Call UpdateDGVs("Studies")

        Call FillUserNames("Studies")

        If boolGuWuOracle Then
            'Try
            '    ta_TBLGUWUSTUDIES.Update(tblGuWuStudies)
            'Catch ex As DBConcurrencyException
            '    'ds2005.tblGuWuProjects.Merge('ds2005.tblGuWuProjects, True)
            'End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_TBLGUWUSTUDIESAcc.Update(tblGuWuStudies)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLGUWUSTUDIES.Merge('ds2005Acc.TBLGUWUSTUDIES, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_TBLGUWUSTUDIESSQLServer.Update(tblGuWuStudies)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLGUWUSTUDIES.Merge('ds2005Acc.TBLGUWUSTUDIES, True)
            End Try
        End If

    End Sub

    Sub SaveAssays()

        Me.dgvAss.CommitEdit(DataGridViewDataErrorContexts.Commit)

        Call UpdateDGVs("GroupDetails")

        If boolGuWuOracle Then
            'Try
            '    ta_TBLGUWUPKSUBJECTS.Update(tblGuWuPKSubjects)
            'Catch ex As DBConcurrencyException
            '    'ds2005Acc.TBLGUWUPKSUBJECTS.Merge('ds2005Acc.TBLGUWUPKSUBJECTS, True)
            'End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_TBLGUWUPKSUBJECTSAcc.Update(tblGuWuPKSubjects)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLGUWUPKSUBJECTS.Merge('ds2005Acc.TBLGUWUPKSUBJECTS, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_TBLGUWUPKSUBJECTSSQLServer.Update(tblGuWuPKSubjects)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLGUWUPKSUBJECTS.Merge('ds2005Acc.TBLGUWUPKSUBJECTS, True)
            End Try
        End If
        If boolGuWuOracle Then
            'Try
            '    ta_TBLGUWURTTIMEPOINTS.Update(tblGuWuRTTimePoints)
            'Catch ex As DBConcurrencyException
            '    'ds2005Acc.TBLGUWURTTIMEPOINTS.Merge('ds2005Acc.TBLGUWURTTIMEPOINTS, True)
            'End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_TBLGUWURTTIMEPOINTSAcc.Update(tblGuWuRTTimePoints)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLGUWURTTIMEPOINTS.Merge('ds2005Acc.TBLGUWURTTIMEPOINTS, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_TBLGUWURTTIMEPOINTSSQLServer.Update(tblGuWuRTTimePoints)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLGUWURTTIMEPOINTS.Merge('ds2005Acc.TBLGUWURTTIMEPOINTS, True)
            End Try
        End If


        If boolGuWuOracle Then
            'Try
            '    ta_TBLGUWUPKROUTES.Update(tblGuWuPKRoutes)
            'Catch ex As DBConcurrencyException
            '    'ds2005Acc.TBLGUWUPKROUTES.Merge('ds2005Acc.TBLGUWUPKROUTES, True)
            'End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_TBLGUWUPKROUTESAcc.Update(tblGuWuPKRoutes)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLGUWUPKROUTES.Merge('ds2005Acc.TBLGUWUPKROUTES, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_TBLGUWUPKROUTESSQLServer.Update(tblGuWuPKRoutes)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLGUWUPKROUTES.Merge('ds2005Acc.TBLGUWUPKROUTES, True)
            End Try
        End If

        If boolGuWuOracle Then
            'Try
            '    ta_TBLGUWUPKGROUPS.Update(tblGuWuPKGroups)
            'Catch ex As DBConcurrencyException
            '    'ds2005Acc.TBLGUWUPKGROUPS.Merge('ds2005Acc.TBLGUWUPKGROUPS, True)
            'End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_TBLGUWUPKGROUPSAcc.Update(tblGuWuPKGroups)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLGUWUPKGROUPS.Merge('ds2005Acc.TBLGUWUPKGROUPS, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_TBLGUWUPKGROUPSSQLServer.Update(tblGuWuPKGroups)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLGUWUPKGROUPS.Merge('ds2005Acc.TBLGUWUPKGROUPS, True)
            End Try
        End If

        If boolGuWuOracle Then
            'Try
            '    ta_TBLGUWUASSIGNEDCMPDLOT.Update(tblGuWuAssignedCmpdLot)
            'Catch ex As DBConcurrencyException
            '    'ds2005Acc.TBLGUWUASSIGNEDCMPDLOT.Merge('ds2005Acc.TBLGUWUASSIGNEDCMPDLOT, True)
            'End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_TBLGUWUASSIGNEDCMPDLOTAcc.Update(tblGuWuAssignedCmpdLot)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLGUWUASSIGNEDCMPDLOT.Merge('ds2005Acc.TBLGUWUASSIGNEDCMPDLOT, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_TBLGUWUASSIGNEDCMPDLOTSQLServer.Update(tblGuWuAssignedCmpdLot)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLGUWUASSIGNEDCMPDLOT.Merge('ds2005Acc.TBLGUWUASSIGNEDCMPDLOT, True)
            End Try
        End If

        If boolGuWuOracle Then
            'Try
            '    ta_TBLGUWUASSIGNEDCMPD.Update(tblGuWuAssignedCmpd)
            'Catch ex As DBConcurrencyException
            '    'ds2005Acc.TBLGUWUASSIGNEDCMPD.Merge('ds2005Acc.TBLGUWUASSIGNEDCMPD, True)
            'End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_TBLGUWUASSIGNEDCMPDAcc.Update(tblGuWuAssignedCmpd)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLGUWUASSIGNEDCMPD.Merge('ds2005Acc.TBLGUWUASSIGNEDCMPD, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_TBLGUWUASSIGNEDCMPDSQLServer.Update(tblGuWuAssignedCmpd)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLGUWUASSIGNEDCMPD.Merge('ds2005Acc.TBLGUWUASSIGNEDCMPD, True)
            End Try
        End If

        If boolGuWuOracle Then
            'Try
            '    ta_TBLGUWUASSAYPERS.Update(TBLGUWUASSAYPERS)
            'Catch ex As DBConcurrencyException
            '    'ds2005Acc.TBLGUWUASSAYPERS.Merge('ds2005Acc.TBLGUWUASSAYPERS, True)
            'End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_TBLGUWUASSAYPERSAcc.Update(TBLGUWUASSAYPERS)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLGUWUASSAYPERS.Merge('ds2005Acc.TBLGUWUASSAYPERS, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_TBLGUWUASSAYPERSSQLServer.Update(TBLGUWUASSAYPERS)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLGUWUASSAYPERS.Merge('ds2005Acc.TBLGUWUASSAYPERS, True)
            End Try
        End If

        Call UpdateDGVs("Assays")

        Call FillUserNames("Assays")

        If boolGuWuOracle Then
            'Try
            '    ta_TBLGUWUASSAY.Update(tblGuWuAssay)
            'Catch ex As DBConcurrencyException
            '    'ds2005Acc.TBLGUWUASSAY.Merge('ds2005Acc.TBLGUWUASSAY, True)
            'End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_TBLGUWUASSAYAcc.Update(tblGuWuAssay)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLGUWUASSAY.Merge('ds2005Acc.TBLGUWUASSAY, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_TBLGUWUASSAYSQLServer.Update(tblGuWuAssay)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLGUWUASSAY.Merge('ds2005Acc.TBLGUWUASSAY, True)
            End Try
        End If

    End Sub

    Sub UpdateDGVs(ByVal strMod As String)

        Dim intRows As Short
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim intRow As Short
        Dim id As Int64
        Dim strF As String
        Dim rows() As DataRow
        Dim dtbl As System.Data.Datatable
        Dim dgvD As DataGridView
        Dim dgvS As DataGridView

        Select Case strMod
            Case "Projects"
                str1 = "ID_TBLGUWUPROJECTS"
                dtbl = tblGuWuProjects
                dgvD = Me.dgvSDProject
                dgvS = Me.dgvProj
            Case "Studies"
                str1 = "ID_TBLGUWUSTUDIES"
                dtbl = tblGuWuStudies
                dgvD = Me.dgvSDStudy
                dgvS = Me.dgvStud
            Case "Assays"
                str1 = "ID_TBLGUWUASSAY"
                dtbl = tblGuWuAssay
                dgvD = Me.dgvAssays
                dgvS = Me.dgvAss
            Case "GroupDetails"
                str1 = "ID_TBLGUWUPKROUTES"
                dtbl = tblGuWuPKRoutes
                dgvD = Me.dgvRoutes
                dgvS = Me.dgvGroupDetails

        End Select

        If dgvD.Rows.Count = 0 Then
            GoTo end1
        End If

        intRow = dgvD.CurrentRow.Index
        id = dgvD.Item(str1, intRow).Value
        strF = str1 & " = " & id
        rows = dtbl.Select(strF)
        intRows = dgvS.Rows.Count

        Dim var1, var2
        Dim int1 As Short

        int1 = 0
        For Count1 = 0 To intRows - 1
            str2 = dgvS.Item("ColumnName", Count1).Value
            str3 = dgvS.Item("ColumnDataType", Count1).Value

            'datatypes
            'system.String
            'system.DateTime
            'system.Int16
            If InStr(str3, "string", CompareMethod.Text) > 0 Then
                var1 = NZ(dgvS.Item("ColumnValueActual", Count1).Value, "")
                var2 = NZ(rows(0).Item(str2), "")
            ElseIf InStr(str3, "datetime", CompareMethod.Text) > 0 Then
                var1 = CDate(NZ(dgvS.Item("ColumnValueActual", Count1).Value, CDate("1/1/1961")))
                var2 = CDate(NZ(rows(0).Item(str2), CDate("1/1/1961")))
            ElseIf InStr(str3, "int", CompareMethod.Text) > 0 Then
                var1 = NZ(dgvS.Item("ColumnValueActual", Count1).Value, 999999999)
                var2 = NZ(rows(0).Item(str2), 999999999)
            End If
            'var1 = dgvS.Item("ColumnValueActual", Count1).Value
            'var2 = rows(0).Item(str2)
            If var1 = var2 Then
            Else
                int1 = int1 + 1
                If int1 = 1 Then
                    rows(0).BeginEdit()
                End If
                Try
                    rows(0).Item(str2) = dgvS.Item("ColumnValueActual", Count1).Value
                Catch ex As Exception

                End Try
            End If
        Next
        If int1 = 0 Then
        Else
            rows(0).EndEdit()
        End If

end1:

    End Sub

    Sub FilterEditUndo()

        Dim dgv As DataGridView
        Dim intRow As Short
        Dim id As Long
        Dim strF As String
        Dim dv As system.data.dataview
        Dim Count1 As Short

        dgv = Me.dgvSDProject
        dgv.Enabled = True
        dgv.DefaultCellStyle.BackColor = Color.White
        dgv = Me.dgvSDProjectS
        dgv.Enabled = True
        dgv.DefaultCellStyle.BackColor = Color.White
        dgv = Me.dgvSDStudy
        dgv.Enabled = True
        dgv.DefaultCellStyle.BackColor = Color.White
        dgv = Me.dgvAssays
        dgv.Enabled = True
        dgv.DefaultCellStyle.BackColor = Color.White

    End Sub

    Sub FilterEdit()

        'in order to force user to edit only one project/study at a time
        'filter project and study dgvs

        'in order to force user to edit only one project/study at a time
        'lock dgv

        Dim dgv As DataGridView
        Dim intRow As Short
        Dim id As Long
        Dim strF As String
        Dim dv As system.data.dataview

        dgv = Me.dgvSDProject
        dgv.Enabled = False
        dgv.DefaultCellStyle.BackColor = Color.Gray
        dgv = Me.dgvSDProjectS
        dgv.Enabled = False
        dgv.DefaultCellStyle.BackColor = Color.Gray
        dgv = Me.dgvSDStudy
        dgv.Enabled = False
        dgv.DefaultCellStyle.BackColor = Color.Gray
        dgv = Me.dgvAssays
        dgv.Enabled = False
        dgv.DefaultCellStyle.BackColor = Color.Gray


    End Sub

    Sub LockCPTab(ByVal bool)

        Me.cmdCPAdd.Enabled = Not (bool)
        Me.cmdCPDelete.Enabled = Not (bool)
        Me.dgvContributingPersonnel.ReadOnly = bool

    End Sub

    Sub SDDataSourceChecked()
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim int1 As Short
        Dim int2 As Short
        Dim str1 As String
        Dim boolHT As Boolean
        Dim strS As String

        boolHT = boolHold
        boolHold = True

        If Me.rbGuWu.Checked Then
            dgv = Me.dgvSDStudy
            int1 = tblGuWuStudies.Rows.Count
            str1 = "CHARSTUDYNAME"
        Else
            dgv = Me.dgvwStudy
            int1 = tblwSTUDY.Rows.Count
            str1 = "STUDYNAME"
        End If

        Try

            If Me.rbGuWu.Checked Then
                dv = New DataView(tblGuWuStudies)
            Else
                dv = New DataView(tblwSTUDY)
            End If

            dv.AllowDelete = False
            dv.AllowEdit = False
            dv.AllowNew = False

            strS = str1 & " ASC"
            dv.Sort = strS

            'assign this datasource to cbxstudy
            'set selection to nothing 
            cbxStudy.DataSource = dv
            Me.cbxStudy.DisplayMember = str1
            Me.cbxStudy.SelectedIndex = -1
            Me.cbxStudy.AutoCompleteMode = AutoCompleteMode.SuggestAppend

            If Me.cbxStudy.Items.Count = 0 Then
                Me.cbxStudy.Text = ""
            End If

        Catch ex As Exception

        End Try

        boolHold = boolHT

    End Sub

    Sub DoThis(ByVal cmd As String)
        Cursor.Current = Cursors.WaitCursor
        Dim int1 As Short
        Dim Count1 As Short
        Dim str1 As String
        Dim strF As String
        Dim bool As Boolean
        Dim boolA As Short

        strF = "ID_TBLPERMISSIONS = " & id_tblPermissions
        Dim rows() As DataRow
        rows = tblPermissions.Select(strF)

        If StrComp(cmd, "Logoff", CompareMethod.Text) = 0 Then
        Else
            If rows.Length = 0 And boolRefresh = False Then
                MsgBox("Guest does not have Edit privileges.", MsgBoxStyle.Information, "No no...")
                Exit Sub
            End If
        End If

        Select Case cmd
            Case "cmdEdit"

                Me.gbxSource.Enabled = False
                Me.cbxStudy.Enabled = False
                Call FilterEdit()

                boolA = BOOLSDPROJECTS
                If boolA = 0 Then
                    bool = False
                Else
                    bool = True
                End If
                Call LockSDProjects(Not (bool))

                boolA = BOOLSDSTUDIES
                If boolA = 0 Then
                    bool = False
                Else
                    bool = True
                End If
                Call LockSDStudies(Not (bool))

                boolA = BOOLCONTRIBUTINGPERSONNEL
                If boolA = 0 Then
                    bool = False
                Else
                    bool = True
                End If
                Call LockCPTab(Not (bool))

                boolA = rows(0).Item("BOOLSDASSAYS")
                If boolA = -1 Then
                    bool = True
                Else
                    bool = False
                End If
                Call LockAssays(Not (bool))

                frmSD.cmdEdit.Enabled = False
                frmSD.cmdSave.Enabled = True
                frmSD.cmdCancel.Enabled = True
                frmSD.cmdExit.Enabled = False

                Call dgvProjectsReadOnly()
                Call dgvStudiesReadOnly()
                Call dgvAssayReadOnly()

                Call CheckStudyRows()

            Case "cmdSave"

                dtG = Now

                Me.gbxSource.Enabled = True
                Me.cbxStudy.Enabled = True

                boolA = BOOLSDPROJECTS
                If boolA = 0 Then
                    bool = False
                Else
                    bool = True
                End If
                If bool Or boolOR Then
                    Call SaveSDProjects()
                End If

                'Call frmsd.ShowThis("SaveSummaryData")

                boolA = BOOLSDSTUDIES
                If boolA = 0 Then
                    bool = False
                Else
                    bool = True
                End If
                If bool Then

                    Call SaveSDStudies()
                End If

                boolA = BOOLCONTRIBUTINGPERSONNEL
                If boolA = 0 Then
                    bool = False
                Else
                    bool = True
                End If
                If bool Or boolOR Then
                    Call SaveCP()
                End If

                boolA = BOOLSDASSAYS
                If boolA = 0 Then
                    bool = False
                Else
                    bool = True
                End If
                If bool Or boolOR Then
                    Call SaveAssays()
                End If

                Call LockAll(True)

                frmSD.cmdEdit.Enabled = True
                frmSD.cmdSave.Enabled = False
                frmSD.cmdCancel.Enabled = False
                frmSD.cmdExit.Enabled = True

                Call FilterEditUndo()

            Case "cmdCancel"
                Me.gbxSource.Enabled = True
                Me.cbxStudy.Enabled = True

                Call DoSDStudiesCancel(False)
                Call DoSDProjectsCancel(False)
                Call DoSDCPCancel(False)
                Call DoAssaysCancel(False)

                Call LockAll(True)

                frmSD.cmdEdit.Enabled = True
                frmSD.cmdSave.Enabled = False
                frmSD.cmdCancel.Enabled = False
                frmSD.cmdExit.Enabled = True

                Call FilterEditUndo()

            Case "cmdExit"

                Call LockAll(True)
                Me.cbxStudy.Enabled = True

                frmSD.cmdEdit.Enabled = True
                frmSD.cmdSave.Enabled = False
                frmSD.cmdCancel.Enabled = False
                frmSD.cmdExit.Enabled = True

        End Select

        boolSDProjAdd = False
        boolSDStudyAdd = False
        SDProjAddID = 0
        SDStudyAddID = 0

        Me.mCal1.Visible = False

        Cursor.Current = Cursors.Default

    End Sub

    Sub DoSDCPCancel(ByVal bool As Boolean)

        Dim tbl As System.Data.Datatable

        Me.dgvContributingPersonnel.CancelEdit()

        tbl = tblContributingPersonnel
        tbl.RejectChanges()

    End Sub

    Sub DoSDProjectsCancel(ByVal bool As Boolean)

        Dim tbl As System.Data.Datatable

        Me.dgvProj.CancelEdit()

        tbl = tblGuWuProjects
        tbl.RejectChanges()

        Call Filldgv1(Me.dgvProj, Me.dgvSDProject, "Projects")


    End Sub

    Sub DoSDStudiesCancel(ByVal bool As Boolean)

        Dim tbl As System.Data.Datatable

        Me.dgvStud.CancelEdit()

        tbl = tblGuWuStudies
        tbl.RejectChanges()

        Call Filldgv1(Me.dgvStud, Me.dgvSDStudy, "Studies")


    End Sub

    Sub DoAssaysCancel(ByVal bool As Boolean)

        boolAssayCancel = True

        Dim tbl As System.Data.Datatable

        Me.dgvAss.CancelEdit()

        tbl = tblGuWuPKSubjects
        tbl.RejectChanges()

        tbl = tblGuWuRTTimePoints
        tbl.RejectChanges()

        tbl = tblGuWuPKRoutes
        tbl.RejectChanges()

        tbl = tblGuWuPKGroups
        tbl.RejectChanges()

        tbl = tblGuWuAssay
        tbl.RejectChanges()

        tbl = tblGuWuAssignedCmpdLot
        tbl.RejectChanges()

        tbl = tblGuWuAssignedCmpd
        tbl.RejectChanges()

        tbl = tblGuWuAssayPERS
        tbl.RejectChanges()

        boolAssayCancel = False

        Call ChangeAssays()

        'Call Filldgv1(Me.dgvAss, Me.dgvAssays, "Assays")

        'Call Filldgv1(Me.dgvGroupDetails, Me.dgvRoutes, "GroupDetails")

        ''Call ChangeCmpds()

        ''Call ChangeAssayPersonnel()

        ''Call UpdateDGVs("Assays")

        'boolAssayCancel = False

        'Call CreateGroupSummary()

    End Sub

    Sub SelectTab1(ByVal intFrom As Short)

        Dim arr1(1)
        Dim intRows As Short
        Dim Count1 As Short
        Dim int1 As Short
        Dim int2 As Short

        intRows = 5
        ReDim arr1(intRows)

        'arr1 is tab1. It's value is sst1 tab
        For Count1 = 0 To intRows
            arr1(Count1) = Count1
        Next

        Try
            If intFrom = 1 Then 'from tab selection
                int1 = Me.dgvTab1.CurrentRow.Index
                Me.sst1.SelectedTab = Me.sst1.TabPages.Item(int1)

            Else
                int2 = Me.sst1.SelectedIndex
                For Count1 = 0 To intRows ' - 1
                    int1 = arr1(Count1)
                    If int1 = int2 Then
                        Try
                            Me.dgvTab1.CurrentCell = Me.dgvTab1.Rows(int1).Cells("CHARITEM")
                        Catch ex As Exception
                            Exit For
                        End Try
                        Exit For
                    End If
                Next
            End If

        Catch ex As Exception

        End Try



    End Sub

    Private Sub dgvwProjectS_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvwProjectS.SelectionChanged

        If boolFormLoad Then
            Exit Sub
        End If

    End Sub

    Private Sub dgvwProject_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvwProject.SelectionChanged

        If boolFormLoad Then
            Exit Sub
        End If

    End Sub

    Private Sub dgvSDProject_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSDProject.SelectionChanged

        If boolFormLoad Or boolHold Then
            Exit Sub
        End If

        Dim intRow As Short

        'intRow = Me.dgvSDProject.CurrentRow.Index
        Try
            intRow = Me.dgvSDProject.CurrentRow.Index
            Try
                Me.dgvSDProjectS.CurrentCell = Me.dgvSDProjectS.Rows.Item(intRow).Cells("CHARPROJECTNAME")
            Catch ex As Exception

            End Try

        Catch ex As Exception

        End Try

        Call Filldgv1(Me.dgvProj, Me.dgvSDProject, "Projects")


    End Sub

    Private Sub dgvSDProjectS_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSDProjectS.SelectionChanged

        Dim strF As String
        Dim boolA As Boolean
        Dim bool As Boolean

        If boolFormLoad Or boolHold Then
            Exit Sub
        End If

        'Call ChangeStudies()
        Call ConfigStudyTableSDdv()

        If Me.dgvSDStudy.Rows.Count = 0 Then
            Call ChangeStudies()
        End If

    End Sub

    Sub ChangeStudies()

        If Me.dgvSDStudy.Rows.Count = 0 Then
            'configure Study stuff
            Call Filldgv1(Me.dgvStud, Me.dgvSDStudy, "Studies")
            id_tblGuWuStudies = -1
        Else


        End If

        'choose cbxStudy
        Call ChoosecbxStudy()

        'configure Personnel stuff

        Call ChangePersonnel()

        Call ChangeAssays()

        'Call ChangeGroups()

        'Call ChangeRoutes()

        'Call ChangeCmpds()

        'Call ChangeAssayPersonnel()


        'pesky
        Call ChoosecbxStudy()

    End Sub

    Sub ChangeAssayPersonnel()
        Dim id As Int64
        Dim intRow As Short
        Dim strF As String
        Dim dv As system.data.dataview
        Dim dgv As DataGridView

        Try
            dv = Me.dgvPI.DataSource
            dgv = Me.dgvAssays
            If dgv.Rows.Count = 0 Then
                id = -1
            ElseIf dgv.CurrentRow Is Nothing Then
                id = dgv.Item("ID_TBLGUWUASSAY", 0).Value
            Else
                intRow = dgv.CurrentRow.Index
                id = dgv.Item("ID_TBLGUWUASSAY", intRow).Value
            End If
            strF = "ID_TBLGUWUASSAY = " & id & " AND CHARROLE = 'PI' AND DTREMOVED IS NULL"
            dv.RowFilter = strF

            Try
                'now fill personnel
                Call FillPersonnel(dv)
            Catch ex As Exception

            End Try

            Call DoLabel(Me.dgvPI, Me.lblPI, "PI's")


        Catch ex As Exception

        End Try

        Try
            dv = Me.dgvAnalyst.DataSource
            dgv = Me.dgvAssays
            If dgv.Rows.Count = 0 Then
                id = -1
            ElseIf dgv.CurrentRow Is Nothing Then
                id = dgv.Item("ID_TBLGUWUASSAY", 0).Value
            Else
                intRow = dgv.CurrentRow.Index
                id = dgv.Item("ID_TBLGUWUASSAY", intRow).Value
            End If
            strF = "ID_TBLGUWUASSAY = " & id & " AND CHARROLE = 'Analyst' AND DTREMOVED IS NULL"
            dv.RowFilter = strF

            Try
                'now fill personnel
                Call FillPersonnel(dv)
            Catch ex As Exception

            End Try

            Call DoLabel(Me.dgvAnalyst, Me.lblAnalyst, "Analysts")


        Catch ex As Exception

        End Try

    End Sub

    Sub ChangeCmpdLots()
        Dim idA As Int64
        Dim id1 As Int64
        Dim intRow As Short
        Dim strF As String
        Dim dv As system.data.dataview
        Dim dgv As DataGridView

        dgv = Me.dgvAssays
        If dgv.Rows.Count = 0 Then
            idA = -1
        ElseIf dgv.CurrentRow Is Nothing Then
            idA = dgv.Item("ID_TBLGUWUASSAY", 0).Value
        Else
            intRow = dgv.CurrentRow.Index
            idA = dgv.Item("ID_TBLGUWUASSAY", intRow).Value
        End If

        Try
            dv = Me.dgvLotNum.DataSource
            dgv = Me.dgvCmpd
            If dgv.Rows.Count = 0 Then
                id1 = -1
            ElseIf dgv.CurrentRow Is Nothing Then
                id1 = dgv.Item("ID_TBLGUWUCOMPOUNDS", 0).Value
            Else
                intRow = dgv.CurrentRow.Index
                'var1 = dgv.Item("ID_TBLGUWUCOMP0UNDS", intRow).Value
                'id1 = dgv.Item("ID_TBLGUWUCOMP0UNDS", intRow).Value

                id1 = dgv("ID_TBLGUWUCOMPOUNDS", intRow).Value

            End If
            strF = "ID_TBLGUWUASSAY = " & idA & " AND ID_TBLGUWUCOMPOUNDS = " & id1
            dv.RowFilter = strF

            Try
                'now fill cmpds
                Call FillCmpdLot(dv)
            Catch ex As Exception

            End Try

        Catch ex As Exception


        End Try

    End Sub

    Sub ChangeCmpds()
        Dim idA As Int64
        Dim id1 As Int64
        Dim intRow As Short
        Dim strF As String
        Dim dv As system.data.dataview
        Dim dgv As DataGridView
        Dim intC As Short
        Dim strT As String

        Try
            dv = Me.dgvCmpd.DataSource
            dgv = Me.dgvAssays
            If dgv.Rows.Count = 0 Then
                idA = -1
            ElseIf dgv.CurrentRow Is Nothing Then
                idA = dgv.Item("ID_TBLGUWUASSAY", 0).Value
            Else
                intRow = dgv.CurrentRow.Index
                idA = dgv.Item("ID_TBLGUWUASSAY", intRow).Value
            End If
            strF = "ID_TBLGUWUASSAY = " & idA
            dv.RowFilter = strF

            Try
                'now fill cmpds
                Call FillCmpd(dv)
            Catch ex As Exception

            End Try

        Catch ex As Exception

        End Try

        Call DoLabel(Me.dgvCmpd, Me.lblCmpd, "Compounds")


    End Sub

    Sub DoLabel(ByVal dgv As DataGridView, ByVal lbl As Label, ByVal strT As String)

        Dim intC As Int16
        intC = dgv.RowCount
        strT = strT & " - " & intC
        lbl.Text = strT

    End Sub

    Sub ChangeRoutes()
        Dim id As Int64
        Dim intRow As Short
        Dim strF As String
        Dim dv As system.data.dataview
        Dim dgv As DataGridView

        Try
            dv = Me.dgvRoutes.DataSource
            dgv = Me.dgvGroups
            If dgv.Rows.Count = 0 Then
                id = -1
            ElseIf dgv.CurrentRow Is Nothing Then
                id = dgv.Item("ID_TBLGUWUPKGROUPS", 0).Value
            Else
                intRow = dgv.CurrentRow.Index
                id = dgv.Item("ID_TBLGUWUPKGROUPS", intRow).Value
            End If
            strF = "ID_TBLGUWUPKGROUPS = " & id
            dv.RowFilter = strF

        Catch ex As Exception

        End Try

        Call DoUnitsTransfer()

        Call DoLabel(Me.dgvRoutes, Me.lbldgvRoutes, "Routes")


    End Sub

    Sub ChangeGroups()

        Dim id As Int64
        Dim intRow As Short
        Dim strF As String
        Dim dv As System.Data.DataView
        Dim dgv As DataGridView

        Try
            dv = Me.dgvGroups.DataSource
            dgv = Me.dgvAssays
            If dgv.Rows.Count = 0 Then
                id = -1
            ElseIf dgv.CurrentRow Is Nothing Then
                id = dgv.Item("ID_TBLGUWUASSAY", 0).Value
            Else
                intRow = dgv.CurrentRow.Index
                id = dgv.Item("ID_TBLGUWUASSAY", intRow).Value
            End If
            strF = "ID_TBLGUWUASSAY = " & id
            dv.RowFilter = strF

            'Call DoLabel(Me.dgvGroups, Me.lbldgvGroups, "Groups")
            Call DoLabel(Me.dgvGroups, Me.lbldgvGroups, "intGroups")


        Catch ex As Exception

        End Try
    End Sub

    Sub ChangeAssays()

        Dim id As Int64
        Dim intRow As Int16
        Dim strF As String
        Dim dv As system.data.dataview


        Try
            dv = Me.dgvAssays.DataSource
            strF = "ID_TBLGUWUSTUDIES = " & id_tblGuWuStudies
            dv.RowFilter = strF

        Catch ex As Exception

        End Try

        Call DoLabel(Me.dgvAssays, Me.lbldgvAssays, "Assays")

        Call Filldgv1(Me.dgvAss, Me.dgvAssays, "Assays")

        Call UpdateDGVs("Assays")

        Call ChangeGroups()

        Call ChangeRoutes()

        Call CreateGroupSummary()

        Call ChangeTimePoints()

        Call ChangePatients()

        Call ChangeCmpds()

        Call ChangeAssayPersonnel()

    End Sub

    Sub ChangePersonnel()

        Dim id As Int64
        Dim intRow As Int16
        Dim strF As String
        Dim dv As system.data.dataview


        Try
            dv = Me.dgvContributingPersonnel.DataSource
            strF = "ID_TBLGUWUSTUDIES = " & id_tblGuWuStudies
            dv.RowFilter = strF

        Catch ex As Exception

        End Try


    End Sub

    Sub CheckStudyRows()
        Dim strF As String
        Dim boolA As Boolean
        Dim bool As Boolean

        If Me.dgvSDStudy.Rows.Count = 0 Then
            Me.dgvStud.ReadOnly = True
        Else

            strF = "ID_TBLPERMISSIONS = " & id_tblPermissions
            Dim rows() As DataRow
            rows = tblPermissions.Select(strF)

            If rows.Length = 0 And boolRefresh = False Then
                bool = False
            Else
                boolA = rows(0).Item("BOOLSDSTUDIES")
                If boolA = -1 Then
                    bool = True
                Else
                    bool = False
                End If
            End If

            If bool Then
                Me.dgvStud.ReadOnly = False
            End If

        End If


    End Sub

    Private Sub cmdEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        Call DoThis("cmdEdit")
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Call DoThis("cmdCancel")

    End Sub

    Private Sub cmdSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Call DoThis("cmdSave")
    End Sub

    Sub FillUserNames(ByVal strMod As String)
        Dim dt As Date
        Dim intRow As Short
        Dim dgv As DataGridView
        Dim rows() As DataRow
        Dim Count1 As Short
        Dim Count2 As Short
        Dim dtbl As System.Data.Datatable
        Dim id As Int64
        Dim strF As String
        Dim strS As String
        Dim boolGo As Boolean = False
        Dim int1 As Short
        Dim int2 As Short
        Dim var1, var2, var3, var4
        Dim str1 As String

        dt = dtG

        Select Case strMod
            Case "Projects"
                If boolSDProjAdd Then

                    dgv = Me.dgvSDProject
                    dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)

                    dtbl = tblGuWuProjects
                    Dim dv As system.data.dataview = New DataView(dtbl)
                    dv.RowStateFilter = DataViewRowState.Added
                    Count1 = dv.Count

                    If Count1 > 0 Then
                        dv(0).BeginEdit()
                        dv(0).Item("CHARUSERIDINIT") = gUserID
                        dv(0).Item("CHARUSERNAMEINIT") = gUserName
                        dv(0).Item("DTINIT") = dt
                        dv(0).Item("CHARUSERIDMOD") = gUserID
                        dv(0).Item("CHARUSERNAMEMOD") = gUserName
                        dv(0).Item("DTMOD") = dt
                        dv(0).EndEdit()
                    End If

                Else

                    dgv = Me.dgvSDProject
                    dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)

                    dtbl = tblGuWuProjects
                    Dim dv As system.data.dataview = New DataView(dtbl)
                    dv.RowStateFilter = DataViewRowState.ModifiedCurrent
                    Dim dv1 As system.data.dataview = New DataView(dtbl)
                    dv1.RowStateFilter = DataViewRowState.ModifiedOriginal
                    int1 = dtbl.Columns.Count
                    boolGo = False
                    For Count1 = 0 To dv1.Count - 1
                        For Count2 = 0 To int1 - 1
                            str1 = dtbl.Columns(Count2).ColumnName
                            If IsUserIDStuff(str1) Then
                            Else
                                var1 = dv(Count1).Item(Count2)
                                var2 = dv1(Count1).Item(Count2)
                                If var1.ToString = var2.ToString Then
                                Else
                                    boolGo = True
                                    Exit For
                                End If
                            End If
                        Next
                        If boolGo Then
                            Exit For
                        End If
                    Next
                    If boolGo Then
                        For Count1 = 0 To dv1.Count - 1
                            dv(Count1).BeginEdit()
                            dv(Count1).Item("CHARUSERIDMOD") = gUserID
                            dv(Count1).Item("CHARUSERNAMEMOD") = gUserName
                            dv(Count1).Item("DTMOD") = dt
                            dv(Count1).EndEdit()
                        Next
                    End If
                End If
            Case "Studies"
                If boolSDStudyAdd Then
                    dgv = Me.dgvSDStudy
                    dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)

                    dtbl = tblGuWuStudies


                    Dim dv As system.data.dataview = New DataView(dtbl)
                    dv.RowStateFilter = DataViewRowState.Added
                    Count1 = dv.Count

                    If Count1 > 0 Then
                        dv(0).BeginEdit()
                        dv(0).Item("CHARUSERIDINIT") = gUserID
                        dv(0).Item("CHARUSERNAMEINIT") = gUserName
                        dv(0).Item("DTINIT") = dt
                        dv(0).Item("CHARUSERIDMOD") = gUserID
                        dv(0).Item("CHARUSERNAMEMOD") = gUserName
                        dv(0).Item("DTMOD") = dt
                        dv(0).EndEdit()
                    End If

                Else
                    dgv = Me.dgvSDStudy
                    dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)

                    dtbl = tblGuWuStudies
                    Dim dv As system.data.dataview = New DataView(dtbl)
                    dv.RowStateFilter = DataViewRowState.ModifiedCurrent
                    Dim dv1 As system.data.dataview = New DataView(dtbl)
                    dv1.RowStateFilter = DataViewRowState.ModifiedOriginal
                    int1 = dtbl.Columns.Count
                    boolGo = False
                    For Count1 = 0 To dv1.Count - 1
                        For Count2 = 0 To int1 - 1
                            str1 = dtbl.Columns(Count2).ColumnName
                            If IsUserIDStuff(str1) Then
                            Else
                                var1 = dv(Count1).Item(Count2)
                                var2 = dv1(Count1).Item(Count2)
                                If var1.ToString = var2.ToString Then
                                Else
                                    boolGo = True
                                    Exit For
                                End If
                            End If
                        Next
                        If boolGo Then
                            Exit For
                        End If
                    Next
                    If boolGo Then
                        For Count1 = 0 To dv1.Count - 1
                            dv(Count1).BeginEdit()
                            dv(Count1).Item("CHARUSERIDMOD") = gUserID
                            dv(Count1).Item("CHARUSERNAMEMOD") = gUserName
                            dv(Count1).Item("DTMOD") = dt
                            dv(Count1).EndEdit()
                        Next
                    End If
                End If
        End Select

    End Sub

    Private Sub cmdAddProject_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddProject.Click
        Dim frm As New frmAddSD

        frm.boolFormLoad = True
        frm.tblT = tblGuWuProjects
        frm.strMod = "Projects"
        frm.intC = 1
        'Call frm.Filldgv1(Me.dgvSDProject)
        Call frm.Filldgv1(Me.dgvProj)

        frm.boolFormLoad = False
        frm.ShowDialog()

        frm.Dispose()


    End Sub

    Private Sub cmdAddStudy_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAddStudy.Click
        Dim frm As New frmAddSD

        frm.boolFormLoad = True
        frm.tblT = tblGuWuStudies
        frm.strMod = "Studies"
        frm.intC = 2
        'Call frm.Filldgv1(Me.dgvSDStudy)
        Call frm.Filldgv1(Me.dgvStud)

        frm.boolFormLoad = False
        frm.ShowDialog()

        frm.Dispose()

    End Sub

    Private Sub dgvProj_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvProj.CellClick

        If e.ColumnIndex = 2 Then
        Else
            GoTo end1
        End If

        Call ForceCellFormat(Me.dgvProj)

end1:

    End Sub

    Private Sub dgvProj_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgvProj.CellValidating

        If boolFormLoad Then
            Exit Sub
        End If

        If Me.cmdEdit.Enabled Then
            Exit Sub
        End If

        Dim dgv As DataGridView
        Dim intRow As Short
        Dim strM As String
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim boolBool As Boolean = False
        Dim boolID As Boolean = False
        Dim strDt As String
        Dim strDt1 As String
        Dim strDt2 As String
        Dim intCol As Short

        dgv = Me.dgvProj
        intCol = e.ColumnIndex
        str1 = dgv.Columns(intCol).Name
        If StrComp(str1, "ColumnValue", CompareMethod.Text) = 0 Then
        Else
            Exit Sub
        End If

        intRow = dgv.CurrentRow.Index
        str1 = dgv("ColumnName", intRow).Value
        str3 = dgv("ColumnHeader", intRow).Value

        If StudProjCellVal(Me.dgvProj, "Projects", str3, e.FormattedValue) Then
            e.Cancel = True
            dgv.CurrentCell = dgv.Rows(intRow).Cells("ColumnValue")
            strM = "The field '" & str3 & "' does not allow duplicates." & ChrW(10) & "Please enter a different value."
            MsgBox(strM, MsgBoxStyle.Information, "Duplicates not allowed...")
        End If

        'now transfer data to ColumnValueActual
        If StrComp(Mid(str1, 1, 4), "bool", CompareMethod.Text) = 0 Then
            boolBool = True
        End If

        If StrComp(Mid(str1, 1, 3), "ID_", CompareMethod.Text) = 0 Then
            boolID = True
        End If

        If boolBool Then
            If StrComp(e.FormattedValue, "TRUE", CompareMethod.Text) = 0 Then
                dgv.Item("ColumnValueActual", intRow).Value = -1
            Else
                dgv.Item("ColumnValueActual", intRow).Value = 0
            End If
        ElseIf boolID Then

        ElseIf IsUserIDStuff(str1) Then
            strDt = Mid(str1, 1, 2)
            If StrComp(strDt, "DT", CompareMethod.Text) = 0 Then
                strDt1 = Format(e.FormattedValue, "Long Date")
                strDt2 = Format(e.FormattedValue, "Long Time")
                str1 = strDt1 & " " & strDt2
                dgv.Item("ColumnValueActual", intRow).Value = str1
            Else
                dgv.Item("ColumnValueActual", intRow).Value = e.FormattedValue
            End If

        Else
            dgv.Item("ColumnValueActual", intRow).Value = e.FormattedValue
        End If


    End Sub

    Private Sub dgvStud_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvStud.CellClick

        If e.ColumnIndex = 2 Then
        Else
            GoTo end1
        End If

        Call ForceCellFormat(Me.dgvStud)

end1:
    End Sub

    Private Sub dgvStud_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgvStud.CellValidating

        If boolFormLoad Then
            Exit Sub
        End If

        If Me.cmdEdit.Enabled Then
            Exit Sub
        End If

        Dim dgv As DataGridView
        Dim intRow As Short
        Dim strM As String
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim boolBool As Boolean = False
        Dim boolID As Boolean = False
        Dim strF As String
        Dim rows() As DataRow
        Dim int1 As Int64
        Dim strDt As String
        Dim strDt1 As String
        Dim strDt2 As String
        Dim intCol As Short

        dgv = Me.dgvStud
        intCol = e.ColumnIndex
        str1 = dgv.Columns(intCol).Name
        If StrComp(str1, "ColumnValue", CompareMethod.Text) = 0 Then
        Else
            Exit Sub
        End If

        intRow = dgv.CurrentRow.Index
        str1 = dgv("ColumnName", intRow).Value
        str3 = dgv("ColumnHeader", intRow).Value

        If StudProjCellVal(Me.dgvStud, "Studies", str3, e.FormattedValue) Then
            e.Cancel = True
            dgv.CurrentCell = dgv.Rows(intRow).Cells("ColumnValue")
            strM = "The field '" & str3 & "' does not allow duplicates." & ChrW(10) & "Please enter a different value."
            MsgBox(strM, MsgBoxStyle.Information, "Duplicates not allowed...")
            GoTo end1
        End If

        'now transfer data to ColumnValueActual
        If StrComp(Mid(str1, 1, 4), "bool", CompareMethod.Text) = 0 Then
            boolBool = True
        End If

        If StrComp(Mid(str1, 1, 3), "ID_", CompareMethod.Text) = 0 Then
            boolID = True
        End If

        If boolBool Then
            If StrComp(e.FormattedValue, "TRUE", CompareMethod.Text) = 0 Then
                dgv.Item("ColumnValueActual", intRow).Value = -1
            Else
                dgv.Item("ColumnValueActual", intRow).Value = 0
            End If
        ElseIf boolID Then
            If StrComp(str1, "ID_TBLGUWUPROJECTS", CompareMethod.Text) = 0 Then


            ElseIf StrComp(str1, "ID_TBLCONFIGREPORTTYPE", CompareMethod.Text) = 0 Then
                If Len(NZ(e.FormattedValue, "")) = 0 Then
                    dgv.Item("ColumnValueActual", intRow).Value = DBNull.Value
                Else
                    strF = "CHARREPORTTYPE = '" & e.FormattedValue & "'"
                    rows = tblConfigReportType.Select(strF)

                    int1 = rows(0).Item("ID_TBLCONFIGREPORTTYPE")
                    If int1 = 0 Then
                        dgv.Item("ColumnValueActual", intRow).Value = DBNull.Value
                    Else
                        dgv.Item("ColumnValueActual", intRow).Value = rows(0).Item("ID_TBLCONFIGREPORTTYPE")
                    End If
                End If


            ElseIf StrComp(str1, "ID_TBLGUWUSTUDYSTAT", CompareMethod.Text) = 0 Then
                If Len(NZ(e.FormattedValue, "")) = 0 Then
                    dgv.Item("ColumnValueActual", intRow).Value = DBNull.Value
                Else
                    strF = "CHARSTATUS = '" & e.FormattedValue & "'"
                    rows = tblGuWuStudyStat.Select(strF)
                    int1 = rows(0).Item("ID_TBLGUWUSTUDYSTAT")
                    If int1 = 0 Then
                        dgv.Item("ColumnValueActual", intRow).Value = DBNull.Value
                    Else
                        dgv.Item("ColumnValueActual", intRow).Value = rows(0).Item("ID_TBLGUWUSTUDYSTAT")
                    End If

                End If


            ElseIf StrComp(str1, "ID_TBLGUWUSTUDYDESIGNTYPE", CompareMethod.Text) = 0 Then
                If Len(NZ(e.FormattedValue, "")) = 0 Then
                    dgv.Item("ColumnValueActual", intRow).Value = DBNull.Value
                Else
                    strF = "CHARSTUDYDESIGNTYPE = '" & e.FormattedValue & "'"
                    rows = tblGuWuStudyDesignType.Select(strF)
                    int1 = rows(0).Item("ID_TBLGUWUSTUDYDESIGNTYPE")
                    If int1 = 0 Then
                        dgv.Item("ColumnValueActual", intRow).Value = DBNull.Value
                    Else
                        dgv.Item("ColumnValueActual", intRow).Value = rows(0).Item("ID_TBLGUWUSTUDYDESIGNTYPE")
                    End If

                End If

            End If
        ElseIf IsUserIDStuff(str1) Then
            strDt = Mid(str1, 1, 2)
            If StrComp(strDt, "DT", CompareMethod.Text) = 0 Then
                strDt1 = Format(e.FormattedValue, "Long Date")
                strDt2 = Format(e.FormattedValue, "Long Time")
                str1 = strDt1 & " " & strDt2
                dgv.Item("ColumnValueActual", intRow).Value = str1
            Else
                dgv.Item("ColumnValueActual", intRow).Value = e.FormattedValue
            End If
        Else
            dgv.Item("ColumnValueActual", intRow).Value = e.FormattedValue
        End If

end1:

    End Sub

    Function StudProjCellVal(ByVal dgv As DataGridView, ByVal strMod As String, ByVal str3 As String, ByVal svar1 As Object) As Boolean
        'Dim dgv As DataGridView
        Dim intRow As Short
        Dim var1
        Dim strF As String
        Dim rows() As DataRow
        Dim dtbl As System.Data.Datatable
        Dim str1 As String
        Dim str2 As String
        'Dim str3 As String
        Dim boolGo As Boolean
        Dim boolHit As Boolean
        Dim strM As String
        Dim id As Int64
        Dim intRowID As Short


        'dgv = Me.dgv1
        intRow = dgv.CurrentRow.Index
        str1 = dgv("ColumnName", intRow).Value
        str3 = dgv("ColumnHeader", intRow).Value


        boolGo = False
        boolHit = False
        Select Case strMod
            Case "Projects"
                If Me.dgvSDProject.CurrentRow Is Nothing Then
                    GoTo end1
                End If
                intRowID = Me.dgvSDProject.CurrentRow.Index
                id = Me.dgvSDProject.Item("ID_TBLGUWUPROJECTS", intRowID).Value
            Case "Studies"
                If Me.dgvSDStudy.CurrentRow Is Nothing Then
                    GoTo end1
                End If
                intRowID = Me.dgvSDStudy.CurrentRow.Index
                id = Me.dgvSDStudy.Item("ID_TBLGUWUSTUDIES", intRowID).Value
        End Select
        'var1 = NZ(dgv("ColumnValue", intRow).Value, "")
        'svar1 = NZ(e.FormattedValue, "")


        If StrComp(strMod, "Projects", CompareMethod.Text) = 0 Then
            If StrComp(str1, "CHARPROJECTNAME", CompareMethod.Text) = 0 Then
                boolGo = True
                strF = "CHARPROJECTNAME = '" & svar1 & "'"
            ElseIf StrComp(str1, "CHARPROJECTNUM", CompareMethod.Text) = 0 Then
                boolGo = True
                strF = "CHARPROJECTNUM = '" & svar1 & "'"
            End If
            If boolGo Then
                strF = strF & " AND ID_TBLGUWUPROJECTS <> " & id
                dtbl = tblGuWuProjects
                rows = dtbl.Select(strF)
                If rows.Length = 0 Then 'OK
                Else
                    boolHit = True
                End If
            End If

        ElseIf StrComp(strMod, "Studies", CompareMethod.Text) = 0 Then
            If StrComp(str1, "CHARSTUDYNAME", CompareMethod.Text) = 0 Then
                boolGo = True
                strF = "CHARSTUDYNAME = '" & svar1 & "'"
            ElseIf StrComp(str1, "CHARSTUDYNUM", CompareMethod.Text) = 0 Then
                boolGo = True
                strF = "CHARSTUDYNUM = '" & svar1 & "'"
            End If
            If boolGo Then
                strF = strF & " AND ID_TBLGUWUSTUDIES <> " & id
                dtbl = tblGuWuStudies
                rows = dtbl.Select(strF)
                If rows.Length = 0 Then 'OK
                Else
                    boolHit = True
                End If
            End If

        End If

end1:

        StudProjCellVal = boolHit

    End Function

    Private Sub mCal1_DateSelected(ByVal sender As Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles mCal1.DateSelected

        Dim intRow As Short

        If boolFromWeekRange Then
            If boolLL Then
                Me.txtLL.Text = e.Start

            Else
                Me.txtUL.Text = e.Start
            End If
            boolFromWeekRange = False
            boolLL = False
            boolUL = False
            Me.mCal1.Visible = False
        Else
            intRow = dgvA.CurrentRow.Index

            dgvA.Rows(intRow).Cells("ColumnValue").Value = e.Start

        End If


    End Sub

    Private Sub dgvSDStudy_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSDStudy.SelectionChanged

        If boolFormLoad Then
            Exit Sub
        End If

        Call Filldgv1(Me.dgvStud, Me.dgvSDStudy, "Studies")

        Call ChangeStudies()


    End Sub

    Private Sub rbProjectList_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbProjectList.CheckedChanged
        'Call ShowGuWudgv()
        Call Configdgv1Again("Projects")
        Call ShowGuWudgv()

    End Sub

    Private Sub rbStudyList_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbStudyList.CheckedChanged
        Call Configdgv1Again("Studies")
        Call ShowGuWudgv()

    End Sub

    Sub CPTab_Initialize()
        Dim tbl As System.Data.Datatable
        Dim dtbl As System.Data.Datatable
        Dim dgv As DataGridView
        Dim drow As DataRow
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim Count1 As Short
        Dim boolg As Boolean
        Dim strT As String
        Dim strF As String
        Dim var1, var2
        Dim wi As Short
        Dim twi As Short
        Dim col2 As DataColumn
        Dim col As DataColumn
        Dim strS As String
        Dim boolV As Boolean
        Dim int1 As Short

        tbl = tblCP
        dgv = Me.dgvContributingPersonnel
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader
        dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        dtbl = tblContributingPersonnel

        'add one unbound columns
        If dtbl.Columns.Contains("boolIncludeSOTP") Then
        Else
            Dim col1 As New DataColumn
            str1 = "boolIncludeSOTP"
            str2 = "A_*"
            col1.ColumnName = str1
            col1.Caption = str2
            col1.DataType = System.Type.GetType("System.Boolean")
            col1.DefaultValue = False
            col1.AllowDBNull = False
            dtbl.Columns.Add(col1)

        End If

        'For Count1 = 1 To 3
        '    Dim col1 As New DataColumn
        '    Select Case Count1
        '        Case 1
        '            str1 = "boolIncludeSOP"
        '            str2 = "A_*"
        '        Case 2
        '            str1 = "boolIncludeSOTP"
        '            str2 = "B_*"
        '        Case 3
        '            str1 = "boolIncludeSOCS"
        '            str2 = "C_*"
        '    End Select
        '    col1.ColumnName = str1
        '    col1.Caption = str2
        '    col1.DataType = System.Type.GetType("System.Boolean")
        '    col1.DefaultValue = False
        '    col1.AllowDBNull = False
        '    dtbl.Columns.Add(col1)
        'Next

        strF = "id_tblStudies = " & id_tblStudies
        strS = "intOrder ASC"
        Dim dv As system.data.dataview = New DataView(dtbl, strF, strS, DataViewRowState.CurrentRows)
        dv.AllowDelete = False
        dv.AllowNew = False
        dgv.DataSource = dv

        'make all columns invisible
        For Count1 = 0 To dtbl.Columns.Count - 1
            dgv.Columns.Item(Count1).Visible = False
            dgv.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable

        Next

        Dim strCol As String

        twi = CInt(dgv.Width)
        wi = 0
        int1 = dgv.Columns.Count
        Dim tl As Short
        Dim ctl As Control
        Dim ctl1 As Control
        tl = dgv.RowHeadersWidth + dgv.Left
        For Count1 = 1 To 8
            boolg = False
            boolV = False
            Select Case Count1
                Case 1
                    str3 = "charCPPrefix"
                    boolg = True
                    boolV = True
                    str1 = "text"
                    str2 = "Prefix"
                    wi = twi * 0.05
                Case 2
                    str3 = "charCPName"
                    boolg = True
                    boolV = True
                    str1 = "text"
                    str2 = "Name **"
                    wi = twi * 0.2
                Case 3
                    str3 = "charCPSuffix"
                    boolg = True
                    boolV = True
                    str1 = "text"
                    str2 = "Suffix"
                    wi = twi * 0.05
                Case 4
                    str3 = "charCPDegree"
                    boolg = True
                    boolV = True
                    str1 = "text"
                    str2 = "Degree"
                    wi = twi * 0.06
                Case 5
                    str3 = "charCPTitle"
                    boolg = True
                    boolV = True
                    str1 = "text"
                    str2 = "Title"
                    wi = twi * 0.1675
                Case 6
                    str3 = "charCPRole"
                    boolg = True
                    boolV = True
                    str1 = "text"
                    str2 = "Role"
                    wi = twi * 0.1675
                    'Case 7
                    '    str3 = "boolIncludeSOP"
                    '    boolg = False
                    '    boolV = True
                    '    str1 = "bool"
                    '    str2 = "A *"
                    '    wi = twi * 0.04
                Case 7
                    str3 = "boolIncludeSOTP"
                    boolg = False
                    boolV = False
                    str1 = "bool"
                    str2 = "A *"
                    wi = twi * 0.04
                    'Case 9
                    '    str3 = "boolIncludeSOCS"
                    '    boolg = False
                    '    boolV = True
                    '    str1 = "bool"
                    '    str2 = "C *"
                    '    wi = twi * 0.04
                Case 8
                    str3 = "intOrder"
                    boolg = False
                    boolV = False
                    str1 = "textc"
                    str2 = "B *"
                    wi = twi * 0.04

            End Select

            If boolV Then
                dgv.Columns.Item(str3).Visible = boolV
                dgv.Columns.Item(str3).HeaderText = str2
                dgv.Columns.Item(str3).Width = wi
                dgv.Columns.Item(str3).MinimumWidth = wi
            End If
            If StrComp(str3, "intOrder", CompareMethod.Text) = 0 Then
                dgv.Columns.Item(str3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End If
            'DisplayIndex
        Next

        'set displayorder
        Call CPDisplayOrder()

        Call FillDropdownBoxes()

        Call CP_FillTable()

    End Sub

    Sub CPDisplayOrder()
        Dim dgv As DataGridView
        Dim Count1 As Short
        Dim int1 As Short
        Dim int2 As Short
        Dim str1 As String
        Dim str3 As String
        Dim boolVis As Boolean

        dgv = Me.dgvContributingPersonnel
        int1 = dgv.Columns.Count
        'first set everything to a high display order
        'For Count1 = 0 To int1 - 1
        '    dgv.Columns.Item(Count1).DisplayIndex = int1 - 1
        'Next

        'set display order
        For Count1 = 0 To int1 - 1
            str1 = ""
            str3 = dgv.Columns(Count1).Name
            boolVis = True
            Select Case str3
                Case UCase("charCPPrefix")
                    str1 = "charCPPrefix"
                    int2 = 0
                Case UCase("charCPName")
                    str1 = "charCPName"
                    int2 = 1
                Case UCase("charCPSuffix")
                    str1 = "charCPSuffix"
                    int2 = 2
                Case UCase("charCPDegree")
                    str1 = "charCPDegree"
                    int2 = 3
                Case UCase("charCPTitle")
                    str1 = "charCPTitle"
                    int2 = 4
                Case UCase("charCPRole")
                    str1 = "charCPRole"
                    int2 = 5
                    'Case 7
                    '    str3 = "boolIncludeSOP"
                    '    int2 = 6
                Case UCase("boolIncludeSOTP")
                    str1 = "boolIncludeSOTP"
                    int2 = 6
                    'Case 9
                    '    str3 = "boolIncludeSOCS"
                    '    int2 = 8
                    boolVis = False
                Case UCase("intOrder")
                    str1 = "intOrder"
                    int2 = 7
                    boolVis = False

            End Select
            If Len(str1) = 0 Then
            Else
                dgv.Columns.Item(str3).DisplayIndex = int2
                dgv.Columns.Item(str3).Visible = boolVis
            End If

            ''set display order
            'For Count1 = 0 To int1 - 1
            '    str3 = ""
            '    str3 = dgv.Columns(Count1).Name
            '    Select Case Count1
            '        Case 1
            '            str3 = "charCPPrefix"
            '            int2 = 0
            '        Case 2
            '            str3 = "charCPName"
            '            int2 = 1
            '        Case 3
            '            str3 = "charCPSuffix"
            '            int2 = 2
            '        Case 4
            '            str3 = "charCPDegree"
            '            int2 = 3
            '        Case 5
            '            str3 = "charCPTitle"
            '            int2 = 4
            '        Case 6
            '            str3 = "charCPRole"
            '            int2 = 5
            '            'Case 7
            '            '    str3 = "boolIncludeSOP"
            '            '    int2 = 6
            '        Case 7
            '            str3 = "boolIncludeSOTP"
            '            int2 = 6
            '            'Case 9
            '            '    str3 = "boolIncludeSOCS"
            '            '    int2 = 8
            '        Case 8
            '            str3 = "intOrder"
            '            int2 = 7

            '    End Select
            '    If Len(str3) = 0 Then
            '    Else
            '        dgv.Columns.Item(str3).DisplayIndex = int2
            '    End If
        Next

    End Sub


    Sub FillDropdownBoxes()
        Dim tbl As System.Data.Datatable
        Dim strF As String
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim var1

        boolStopCBX = True

        'retrieve contributing personnel Degree
        Dim rows() As DataRow
        Dim strS As String
        Dim int1 As Short
        tbl = tblDropdownBoxContent
        Erase rows
        strF = "id_tblDropdownBoxName = 4"
        strS = "intOrder ASC"
        rows = tbl.Select(strF, strS)
        cbxxCPDegree.Items.Clear()
        cbxxCPDegree.Items.Add("")
        int1 = rows.Length
        'var1 = frmh.cbxCPDegree.SelectedIndex
        For Count1 = 0 To int1 - 1
            str1 = rows(Count1).Item("charValue")
            cbxxCPDegree.Items.Add(str1)
        Next
        'cbxxCPDegree.Value = ""
        cbxxCPDegree.AutoComplete = True
        cbxxCPDegree.MaxDropDownItems = 20
        cbxxCPDegree.Sorted = True
        cbxxCPDegree.DisplayStyleForCurrentCellOnly = True
        cbxxCPDegree.DropDownWidth = cbxxCPDegree.DropDownWidth * 1.5
        cbxxCPDegree.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton

        'retrieve contributing personnel title
        Erase rows
        strF = "id_tblDropdownBoxName = 8"
        strS = "intOrder ASC"
        rows = tbl.Select(strF, strS)
        cbxxCPTitle.Items.Clear()
        cbxxCPTitle.Items.Add("")
        int1 = rows.Length
        'var1 = frmh.cbxCPTitle.SelectedIndex
        For Count1 = 0 To int1 - 1
            str1 = rows(Count1).Item("charValue")
            cbxxCPTitle.Items.Add(str1)
        Next
        'cbxxCPTitle.Value = ""
        cbxxCPTitle.AutoComplete = True
        cbxxCPTitle.MaxDropDownItems = 20
        cbxxCPTitle.Sorted = True
        cbxxCPTitle.DisplayStyleForCurrentCellOnly = True
        cbxxCPTitle.DropDownWidth = cbxxCPTitle.DropDownWidth * 1.5
        cbxxCPTitle.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton

        'retrieve contributing personnel prefix
        Erase rows
        strF = "id_tblDropdownBoxName = 5"
        strS = "intOrder ASC"
        rows = tbl.Select(strF, strS)
        'frmh.cbxCPPrefix.Items.Add("[None]")
        cbxxCPPrefix.Items.Clear()
        cbxxCPPrefix.Items.Add("")
        int1 = rows.Length
        'var1 = frmh.cbxCPTitle.SelectedIndex
        For Count1 = 0 To int1 - 1
            str1 = rows(Count1).Item("charValue")
            cbxxCPPrefix.Items.Add(str1)
        Next
        'cbxxCPPrefix.Value = ""
        cbxxCPPrefix.AutoComplete = True
        cbxxCPPrefix.MaxDropDownItems = 20
        cbxxCPPrefix.DisplayStyleForCurrentCellOnly = True
        cbxxCPPrefix.DropDownWidth = cbxxCPPrefix.DropDownWidth * 1.5
        cbxxCPPrefix.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton

        'retrieve contributing personnel suffix
        Erase rows
        strF = "id_tblDropdownBoxName = 6"
        strS = "intOrder ASC"
        rows = tbl.Select(strF, strS)
        'frmh.cbxCPSuffix.Items.Add("[None]")
        cbxxCPSuffix.Items.Clear()
        cbxxCPSuffix.Items.Add("")
        int1 = rows.Length
        'var1 = frmh.cbxCPTitle.SelectedIndex
        For Count1 = 0 To int1 - 1
            str1 = rows(Count1).Item("charValue")
            cbxxCPSuffix.Items.Add(str1)
        Next
        'cbxxCPSuffix.Value = ""
        cbxxCPSuffix.AutoComplete = True
        cbxxCPSuffix.MaxDropDownItems = 20
        cbxxCPSuffix.DisplayStyleForCurrentCellOnly = True
        cbxxCPSuffix.DropDownWidth = cbxxCPSuffix.DropDownWidth * 1.5
        cbxxCPSuffix.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton

        'retrieve contributing personnel role
        Erase rows
        strF = "id_tblDropdownBoxName = 7"
        strS = "intOrder ASC"
        rows = tbl.Select(strF, strS)
        'frmh.cbxCPRole.Items.Add("[None]")
        cbxxCPRole.Items.Clear()
        cbxxCPRole.Items.Add("")
        int1 = rows.Length
        'var1 = frmh.cbxCPTitle.SelectedIndex
        For Count1 = 0 To int1 - 1
            str1 = NZ(rows(Count1).Item("charValue"), "")
            'If Len(str1) = 0 Then
            'Else
            '    cbxxCPRole.Items.Add(str1)
            'End If
            cbxxCPRole.Items.Add(str1)
        Next
        'cbxxCPRole.Value = ""
        cbxxCPRole.AutoComplete = True
        cbxxCPRole.MaxDropDownItems = 20
        cbxxCPRole.Sorted = True
        cbxxCPRole.DisplayStyleForCurrentCellOnly = True
        cbxxCPRole.DropDownWidth = cbxxCPRole.DropDownWidth * 1.5
        cbxxCPRole.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton

        'retrieve Assay Proc Descr
        Erase rows
        strF = "id_tblDropdownBoxName = 12"
        strS = "intOrder ASC"
        rows = tbl.Select(strF, strS)
        'frmh.cbxCPRole.Items.Add("[None]")
        cbxxAssayDescr.Items.Clear()
        'cbxxAssayDescr.Items.Add("")
        int1 = rows.Length
        'var1 = frmh.cbxCPTitle.SelectedIndex
        For Count1 = 0 To int1 - 1
            str1 = NZ(rows(Count1).Item("charValue"), "None")
            cbxxAssayDescr.Items.Add(str1)
        Next
        'cbxxCPRole.Value = ""
        cbxxAssayDescr.AutoComplete = True
        cbxxAssayDescr.MaxDropDownItems = 20
        cbxxAssayDescr.Sorted = True
        cbxxAssayDescr.DisplayStyleForCurrentCellOnly = True
        cbxxAssayDescr.DropDownWidth = cbxxCPRole.DropDownWidth * 1.5
        cbxxAssayDescr.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton

        'enter yes/no
        'frmh.cbxCPCoverPageSigBlock.Items.Add("Yes")
        'frmh.cbxCPCoverPageSigBlock.Items.Add("No")
        'frmh.cbxCPCoverPageSigBlock.SelectedIndex = 1
        'frmh.cbxCPTableSigBlock.Items.Add("Yes")
        'frmh.cbxCPTableSigBlock.Items.Add("No")
        'frmh.cbxCPTableSigBlock.SelectedIndex = 1

        'enter Personnel
        tbl = tblPersonnel
        Dim tdv As system.data.dataview = New DataView(tbl)
        'REDO THIS LINE!!!!
        'tdv = tbl.DefaultView
        tdv.Sort = "charLastName ASC"
        int1 = tdv.Count
        cbxxCPName.Items.Clear()
        cbxxCPName.Items.Add("")
        For Count1 = 0 To int1 - 1
            str1 = tdv(Count1).Item("charFIRSTNAME")
            str2 = NZ(tdv(Count1).Item("charMIDDLEname"), "")
            str3 = tdv(Count1).Item("charLASTNAME")
            If StrComp(str3, "aaAdmin", CompareMethod.Text) = 0 Then
            Else
                If Len(str2) = 0 Then 'no middle initial provided
                    str4 = str1 & " " & str3
                Else
                    If Len(str2) = 1 Then 'needs a period
                        str2 = str2 & "."
                    Else
                    End If
                    str4 = str1 & " " & str2 & " " & str3
                End If
                cbxxCPName.Items.Add(str4)
            End If
        Next
        int1 = rows.Length
        'cbxxCPName.Value = ""
        cbxxCPName.AutoComplete = True
        cbxxCPName.MaxDropDownItems = 20
        cbxxCPName.DisplayStyleForCurrentCellOnly = True
        cbxxCPName.DropDownWidth = cbxxCPName.DropDownWidth * 1.5
        cbxxCPName.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton

        boolStopCBX = False

    End Sub

    Sub CP_FillTable()
        Dim dtbl As System.Data.Datatable
        'Dim dv as system.data.dataview
        Dim strF As String
        Dim dgv As DataGridView
        Dim tbl As System.Data.Datatable
        Dim strS As String

        tbl = tblCP
        dgv = Me.dgvContributingPersonnel
        dtbl = tblContributingPersonnel
        dtbl.AcceptChanges()
        strF = "id_tblStudies = " & id_tblStudies
        strS = "intOrder ASC"
        Dim dv As system.data.dataview = New DataView(dtbl, strF, strS, DataViewRowState.CurrentRows)
        dv.AllowNew = False
        dv.AllowEdit = True
        dv.AllowDelete = False
        dgv.DataSource = dv

        Call UpdateCPBool()

        dgv.Refresh()

    End Sub

    Sub UpdateCPBool()
        Dim dv As system.data.dataview
        Dim int1 As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim int2 As Short
        Dim bool As Boolean
        Dim str1 As String
        Dim str2 As String

        dv = Me.dgvContributingPersonnel.DataSource
        int1 = dv.Count
        For Count1 = 0 To int1 - 1
            dv(Count1).BeginEdit()
            str1 = "boolIncludeSOTP"
            str2 = "boolIncludeSigOnTablePage"
            int2 = NZ(dv(Count1).Item(str2), 0)
            If int2 = -1 Then
                bool = True
            Else
                bool = False
            End If
            dv(Count1).Item(str1) = bool
            dv(Count1).EndEdit()

            'For Count2 = 1 To 1
            '    Select Case Count2
            '        Case 1
            '            str1 = "boolIncludeSOP"
            '            str2 = "boolIncludeSigOnCoverPage"
            '        Case 2
            '            str1 = "boolIncludeSOTP"
            '            str2 = "boolIncludeSigOnTablePage"
            '        Case 3
            '            str1 = "boolIncludeSOCS"
            '            str2 = "boolIncludeSigOnCompStatement"
            '    End Select
            '    int2 = dv(Count1).Item(str2)
            '    If int2 = -1 Then
            '        bool = True
            '    Else
            '        bool = False
            '    End If
            '    dv(Count1).Item(str1) = bool
            'Next

            'dv(Count1).EndEdit()
        Next

    End Sub

    Private Sub cmdCPAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCPAdd.Click
        Dim dv As system.data.dataview
        Dim strF As String
        Dim ct1 As Short
        Dim intMax As Short
        Dim Count1 As Short
        Dim int1 As Short
        Dim tbl As System.Data.Datatable
        Dim ct2 As Short
        Dim ct3 As Short
        Dim ct4 As Short
        Dim dtbl As System.Data.Datatable

        Dim tblMax As System.Data.Datatable
        Dim rowsMax() As DataRow
        Dim strFMax As String
        Dim maxID

        maxID = 1

        'strFMax = "charTable = 'tblContributingPersonnel'"
        strFMax = "tblContributingPersonnel"
        maxID = GetMaxID(strFMax, 1, True)


        tbl = tblContributingPersonnel ' tblCP
        ct1 = tbl.Rows.Count

        dv = Me.dgvContributingPersonnel.DataSource
        ct2 = dv.Count

        dv.AllowNew = True
        Dim dvr As DataRowView = dv.AddNew
        dvr.Item("ID_TBLGUWUSTUDIES") = id_tblGuWuStudies
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
        dv.RowFilter = "ID_TBLGUWUSTUDIES = " & id_tblGuWuStudies

        dgvContributingPersonnel.CurrentCell = dgvContributingPersonnel.Rows.Item(ct2).Cells("CHARCPPREFIX")

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

        'If cmdEdit.Enabled Or Len(cbxStudy.Text) = 0 Then
        If cmdEdit.Enabled Then
            Exit Sub
        End If

        If intRow < 0 Or intCol < 0 Then
            Exit Sub
        End If


        dgv = dgvContributingPersonnel
        str1 = dgv.Columns.Item(intCol).Name
        If InStr(1, str1, "bool", CompareMethod.Text) > 0 Then 'ignore
        ElseIf InStr(1, str1, "int", CompareMethod.Text) > 0 Then 'ignore
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
                    var2 = var1 * 1.5
                    cbx.DropDownWidth = var2
                    'if data doesn't exist in dropdown list
                    'data error will be called that inserts unlisted value into dropdown box
                    On Error Resume Next
                    dgv(intCol, intRow) = cbx
                    If Err.Number <> 0 Then
                        Err.Clear()
                    End If
                    On Error GoTo 0
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



    Private Sub cmdCPDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCPDelete.Click
        Dim int1 As Short
        Dim dv As system.data.dataview
        Dim tbl As System.Data.Datatable
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

    Private Sub dgvContributingPersonnel_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvContributingPersonnel.MouseEnter
        Me.dgvContributingPersonnel.Focus()

    End Sub

    Private Sub sst1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles sst1.SelectedIndexChanged
        Call CPDisplayOrder()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Call CreateCalendar()



    End Sub

    Private Sub mnuAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAbout.Click
        Dim frm As New frmAbout

        frm.ShowDialog()

        frm.Dispose()

    End Sub

    Sub ChoosecbxStudy()

        Dim dgv As DataGridView
        Dim boolGuWu As Boolean
        Dim strS As String
        Dim intRow As Int16
        Dim tbl As System.Data.Datatable
        Dim str1 As String
        Dim int1 As Int16
        Dim int2 As Int16
        Dim Count1 As Int16
        Dim strI As String
        Dim varStudyID
        Dim dv As system.data.dataview

        If boolFromcbxStudy Then
            Exit Sub
        End If

        If Me.rbGuWu.Checked Then
            boolGuWu = True
            dgv = Me.dgvSDStudy
            strS = "CHARSTUDYNAME"
            strI = "ID_TBLGUWUSTUDIES"
        Else
            boolGuWu = False
            dgv = Me.dgvwStudy
            strS = "STUDYNAME"
            strI = "STUDYID"
        End If

        If dgv.Rows.Count = 0 Then
            Me.cbxStudy.SelectedIndex = -1
            Exit Sub
        End If

        If dgv.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = dgv.CurrentRow.Index
        End If

        dv = Me.cbxStudy.DataSource

        Try
            'int1 = tbl.Rows.Count
            int1 = dv.Count
            varStudyID = dgv.Item(strI, intRow).Value

            For Count1 = 0 To int1 - 1
                int2 = dv.Item(Count1).Item(strI)
                If int2 = varStudyID Then
                    Me.cbxStudy.SelectedIndex = Count1
                    Exit For
                End If
            Next

        Catch ex As Exception

        End Try

    End Sub

    Private Sub cbxStudy_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxStudy.SelectedIndexChanged
        If boolFormLoad Then
            Exit Sub
        End If

        If boolHold Then
            Exit Sub
        End If

        Call cbxStudyChoose()

    End Sub

    Sub cbxStudyChoose()

        Dim dv As system.data.dataview
        Dim intRow As Int16
        Dim idP As Int64
        Dim idS As Int64
        Dim intRows As Int16
        Dim Count1 As Int16
        Dim int1 As Int64
        Dim boolHT As Boolean

        If boolFromcbxStudy Then
            Exit Sub
        End If

        If Me.rbGuWu.Checked Then
        Else
            Exit Sub
        End If

        dv = Me.cbxStudy.DataSource
        intRow = Me.cbxStudy.SelectedIndex
        boolHT = boolHold

        If intRow = -1 Then
            GoTo end1
        End If

        idP = dv.Item(intRow).Item("ID_TBLGUWUPROJECTS")
        idS = dv.Item(intRow).Item("ID_TBLGUWUSTUDIES")

        boolFromcbxStudy = True

        'set project
        intRows = Me.dgvSDProject.Rows.Count
        For Count1 = 0 To intRows - 1
            int1 = Me.dgvSDProject("ID_TBLGUWUPROJECTS", Count1).Value
            If idP = int1 Then
                Try
                    Me.dgvSDProject.CurrentCell = Me.dgvSDProject.Rows(Count1).Cells("CHARPROJECTNUM")
                Catch ex As Exception

                End Try
                Me.dgvSDProject.Rows(Count1).Selected = True

                'Try
                '    Me.dgvSDProjectS.CurrentCell = Me.dgvSDProjectS.Rows(Count1).Cells("CHARPROJECTNUM")
                'Catch ex As Exception

                'End Try
                'Me.dgvSDProjectS.Rows(Count1).Selected = True

                Exit For
            End If
        Next

        'now do studies
        If Me.dgvSDStudy.Rows.Count = 0 Then
            GoTo end1
        End If

        intRows = Me.dgvSDStudy.Rows.Count
        For Count1 = 0 To intRows - 1
            int1 = Me.dgvSDStudy("ID_TBLGUWUSTUDIES", Count1).Value
            If idS = int1 Then
                Try
                    Me.dgvSDStudy.CurrentCell = Me.dgvSDStudy.Rows(Count1).Cells("CHARSTUDYNUMBER")
                Catch ex As Exception

                End Try
                Me.dgvSDStudy.Rows(Count1).Selected = True
                Exit For
            End If
        Next

end1:

        boolFromcbxStudy = False

    End Sub

    Private Sub dgvAssays_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvAssays.SelectionChanged

        If boolFormLoad Then
            Exit Sub
        End If

        Call ChangeAssays()


        'Call Filldgv1(Me.dgvAss, Me.dgvAssays, "Assays")

        'Call UpdateDGVs("Assays")

        'Call ChangeGroups()

        'Call ChangeRoutes()

        'Call CreateGroupSummary()

        'Call ChangeCmpds()

        'Call ChangeAssayPersonnel()


    End Sub

    Private Sub dgvAss_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvAss.CellClick
        If e.ColumnIndex = 2 Then
        Else
            GoTo end1
        End If

        Call ForceCellFormat(Me.dgvAss)


end1:

    End Sub

    Private Sub dgvAss_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgvAss.CellValidating

        If boolFormLoad Then
            Exit Sub
        End If

        If Me.cmdEdit.Enabled Then
            Exit Sub
        End If

        Dim dgv As DataGridView
        Dim intRow As Short
        Dim strM As String
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim boolBool As Boolean = False
        Dim boolID As Boolean = False
        Dim strF As String
        Dim rows() As DataRow
        Dim int1 As Int64
        Dim strDt As String
        Dim strDt1 As String
        Dim strDt2 As String
        Dim intCol As Short
        Dim strA As String
        Dim strB As String
        Dim intRowA As Short
        Dim var1, var2

        dgv = Me.dgvAss
        intCol = e.ColumnIndex
        str1 = dgv.Columns(intCol).Name
        If StrComp(str1, "ColumnValue", CompareMethod.Text) = 0 Then
        Else
            Exit Sub
        End If

        intRow = dgv.CurrentRow.Index
        str1 = dgv("ColumnName", intRow).Value
        str3 = dgv("ColumnHeader", intRow).Value

        If StudProjCellVal(dgv, "Assays", str3, e.FormattedValue) Then
            e.Cancel = True
            dgv.CurrentCell = dgv.Rows(intRow).Cells("ColumnValue")
            strM = "The field '" & str3 & "' does not allow duplicates." & ChrW(10) & "Please enter a different value."
            MsgBox(strM, MsgBoxStyle.Information, "Duplicates not allowed...")
            GoTo end1
        End If

        'now transfer data to ColumnValueActual
        If StrComp(Mid(str1, 1, 4), "bool", CompareMethod.Text) = 0 Then
            boolBool = True
        End If

        If StrComp(Mid(str1, 1, 3), "ID_", CompareMethod.Text) = 0 Then
            boolID = True
        End If

        If boolBool Then
            If StrComp(e.FormattedValue, "TRUE", CompareMethod.Text) = 0 Then
                dgv.Item("ColumnValueActual", intRow).Value = -1
            Else
                dgv.Item("ColumnValueActual", intRow).Value = 0
            End If
        ElseIf boolID Then
            If StrComp(str1, "ID_TBLGUWUPROJECTS", CompareMethod.Text) = 0 Then


            ElseIf StrComp(str1, "ID_TBLCONFIGREPORTTYPE", CompareMethod.Text) = 0 Then
                If Len(NZ(e.FormattedValue, "")) = 0 Then
                    dgv.Item("ColumnValueActual", intRow).Value = DBNull.Value
                Else
                    strF = "CHARREPORTTYPE = '" & e.FormattedValue & "'"
                    rows = tblConfigReportType.Select(strF)

                    int1 = rows(0).Item("ID_TBLCONFIGREPORTTYPE")
                    If int1 = 0 Then
                        dgv.Item("ColumnValueActual", intRow).Value = DBNull.Value
                    Else
                        dgv.Item("ColumnValueActual", intRow).Value = rows(0).Item("ID_TBLCONFIGREPORTTYPE")
                    End If
                End If


            ElseIf StrComp(str1, "ID_TBLGUWUSTUDYSTAT", CompareMethod.Text) = 0 Then
                If Len(NZ(e.FormattedValue, "")) = 0 Then
                    dgv.Item("ColumnValueActual", intRow).Value = DBNull.Value
                Else
                    strF = "CHARSTATUS = '" & e.FormattedValue & "'"
                    rows = tblGuWuStudyStat.Select(strF)
                    int1 = rows(0).Item("ID_TBLGUWUSTUDYSTAT")
                    If int1 = 0 Then
                        dgv.Item("ColumnValueActual", intRow).Value = DBNull.Value
                    Else
                        dgv.Item("ColumnValueActual", intRow).Value = rows(0).Item("ID_TBLGUWUSTUDYSTAT")
                    End If

                End If


            ElseIf StrComp(str1, "ID_TBLGUWUSTUDYDESIGNTYPE", CompareMethod.Text) = 0 Then
                If Len(NZ(e.FormattedValue, "")) = 0 Then
                    dgv.Item("ColumnValueActual", intRow).Value = DBNull.Value
                Else
                    strF = "CHARSTUDYDESIGNTYPE = '" & e.FormattedValue & "'"
                    rows = tblGuWuStudyDesignType.Select(strF)
                    int1 = rows(0).Item("ID_TBLGUWUSTUDYDESIGNTYPE")
                    If int1 = 0 Then
                        dgv.Item("ColumnValueActual", intRow).Value = DBNull.Value
                    Else
                        dgv.Item("ColumnValueActual", intRow).Value = rows(0).Item("ID_TBLGUWUSTUDYDESIGNTYPE")
                    End If

                End If

            End If
        ElseIf StrComp(str1, "CHARDOSEUNITS", CompareMethod.Text) = 0 Then
            strA = "Target Dose (" & e.FormattedValue & ")"
            intRowA = FindRow(Me.dgvGroupDetails, "CHARTARGETDOSE")
            If intRowA = -1 Then
                strB = "NA"
            Else
                strB = NZ(Me.dgvGroupDetails("ColumnHeader", intRowA).Value, "NA")
            End If
            If StrComp(strA, strB, CompareMethod.Text) = 0 Then
            Else
                If intRowA = -1 Then
                Else
                    Me.dgvGroupDetails("ColumnHeader", intRowA).Value = strA
                    Me.dgvGroupDetails.AutoResizeColumns()
                End If
            End If
            dgv.Item("ColumnValueActual", intRow).Value = e.FormattedValue

        ElseIf StrComp(str1, "CHARDOSECONCUNITS", CompareMethod.Text) = 0 Then
            strA = "Dose Conc. (" & e.FormattedValue & ")"
            intRowA = FindRow(Me.dgvGroupDetails, "CHARTARGETDOSECONC")
            If intRowA = -1 Then
                strB = "NA"
            Else
                strB = NZ(Me.dgvGroupDetails("ColumnHeader", intRowA).Value, "NA")
            End If
            If StrComp(strA, strB, CompareMethod.Text) = 0 Then
            Else
                If intRowA = -1 Then
                Else
                    Me.dgvGroupDetails("ColumnHeader", intRowA).Value = strA
                    Me.dgvGroupDetails.AutoResizeColumns()
                End If
            End If
            dgv.Item("ColumnValueActual", intRow).Value = e.FormattedValue

        ElseIf StrComp(str1, "CHARTISSUEWTUNITS", CompareMethod.Text) = 0 Then
            strA = "Target Tissue Wt. (" & e.FormattedValue & ")"
            intRowA = FindRow(Me.dgvGroupDetails, "CHARTARGETTISSUEWT")
            If intRowA = -1 Then
                strB = "NA"
            Else
                strB = NZ(Me.dgvGroupDetails("ColumnHeader", intRowA).Value, "NA")
            End If
            If StrComp(strA, strB, CompareMethod.Text) = 0 Then
            Else
                If intRowA = -1 Then
                Else
                    Me.dgvGroupDetails("ColumnHeader", intRowA).Value = strA
                    Me.dgvGroupDetails.AutoResizeColumns()
                End If
            End If
            dgv.Item("ColumnValueActual", intRow).Value = e.FormattedValue

        ElseIf IsUserIDStuff(str1) Then
            strDt = Mid(str1, 1, 2)
            If StrComp(strDt, "DT", CompareMethod.Text) = 0 Then
                strDt1 = Format(e.FormattedValue, "Long Date")
                strDt2 = Format(e.FormattedValue, "Long Time")
                str1 = strDt1 & " " & strDt2
                dgv.Item("ColumnValueActual", intRow).Value = str1
            Else
                dgv.Item("ColumnValueActual", intRow).Value = e.FormattedValue
            End If
        Else
            dgv.Item("ColumnValueActual", intRow).Value = e.FormattedValue
        End If

end1:
    End Sub

    Sub DoUnitsTransfer()

        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView
        Dim str1 As String
        Dim strA As String
        Dim strB As String
        Dim intRow As Short
        Dim intRowA As Short
        Dim Count1 As Short
        Dim var1, var2

        dgv1 = Me.dgvAss
        dgv2 = Me.dgvGroupDetails

        'do CHARDOSEUNITS
        For Count1 = 1 To 3
            strA = "NA"
            strB = "NA"
            Select Case Count1
                Case 1
                    str1 = "CHARDOSEUNITS"
                Case 2
                    str1 = "CHARDOSECONCUNITS"
                Case 3
                    str1 = "CHARTISSUEWTUNITS"
            End Select
            If StrComp(str1, "CHARDOSEUNITS", CompareMethod.Text) = 0 Then
                intRow = FindRow(Me.dgvAss, "CHARDOSEUNITS")
                If intRow = -1 Then
                    GoTo next1
                End If
                var1 = Me.dgvAss("ColumnValue", intRow).Value

                strA = "Target Dose (" & var1 & ")"
                intRowA = FindRow(Me.dgvGroupDetails, "CHARTARGETDOSE")
                If intRowA = -1 Then
                    strB = "NA"
                Else
                    strB = NZ(Me.dgvGroupDetails("ColumnHeader", intRowA).Value, "NA")
                End If
                If StrComp(strA, strB, CompareMethod.Text) = 0 Then
                Else
                    If intRowA = -1 Then
                    Else
                        Me.dgvGroupDetails("ColumnHeader", intRowA).Value = strA
                        Me.dgvGroupDetails.AutoResizeColumns()
                    End If
                End If

            ElseIf StrComp(str1, "CHARDOSECONCUNITS", CompareMethod.Text) = 0 Then
                intRow = FindRow(Me.dgvAss, "CHARDOSECONCUNITS")
                If intRow = -1 Then
                    GoTo next1
                End If
                var1 = Me.dgvAss("ColumnValue", intRow).Value

                strA = "Dose Conc. (" & var1 & ")"
                intRowA = FindRow(Me.dgvGroupDetails, "CHARTARGETDOSECONC")
                If intRowA = -1 Then
                    strB = "NA"
                Else
                    strB = NZ(Me.dgvGroupDetails("ColumnHeader", intRowA).Value, "NA")
                End If
                If StrComp(strA, strB, CompareMethod.Text) = 0 Then
                Else
                    If intRowA = -1 Then
                    Else
                        Me.dgvGroupDetails("ColumnHeader", intRowA).Value = strA
                        Me.dgvGroupDetails.AutoResizeColumns()
                    End If
                End If

            ElseIf StrComp(str1, "CHARTISSUEWTUNITS", CompareMethod.Text) = 0 Then
                intRow = FindRow(Me.dgvAss, "CHARTISSUEWTUNITS")
                If intRow = -1 Then
                    GoTo next1
                End If
                var1 = Me.dgvAss("ColumnValue", intRow).Value

                strA = "Target Tissue Wt. (" & var1 & ")"
                intRowA = FindRow(Me.dgvGroupDetails, "CHARTARGETTISSUEWT")
                If intRowA = -1 Then
                    strB = "NA"
                Else
                    strB = NZ(Me.dgvGroupDetails("ColumnHeader", intRowA).Value, "NA")
                End If
                If StrComp(strA, strB, CompareMethod.Text) = 0 Then
                Else
                    If intRowA = -1 Then
                    Else
                        Me.dgvGroupDetails("ColumnHeader", intRowA).Value = strA
                        Me.dgvGroupDetails.AutoResizeColumns()
                    End If
                End If
            End If

next1:
        Next

    End Sub


    Private Sub cmdGetAssay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGetAssay.Click

        Dim strM As String
        Dim strM1 As String
        Dim str1 As String
        Dim intRow As Short
        Dim intGo As Short

        intRow = 0
        If Me.dgvAssays.Rows.Count = 0 Then
            strM = "An Assay must be configured before you can apply a template."
        ElseIf Me.dgvAssays.CurrentRow Is Nothing Then
            strM = "Please select an Assay."
        Else
            strM = ""
            intRow = Me.dgvAssays.CurrentRow.Index
            str1 = Me.dgvAssays("CHARASSAYNAME", intRow).Value
        End If

        If Len(strM) = 0 Then

            strM1 = "Performing this action will overwrite any data that exists in Assay '" & str1 & "'."
            strM1 = strM1 & ChrW(10) & ChrW(10) & "Do you wish to continue?"
            intGo = MsgBox(strM1, MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Do you wish to continue...")
            If intGo = 6 Then 'continue
            Else
                GoTo end1
            End If
        End If

        If Len(strM) = 0 Then
            Call AddAssay(False, True)
        Else
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
        End If

end1:

    End Sub

    Private Sub cmdAddAssay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddAssay.Click

        Call AddAssay(True, False)

        Call DoAllLabels()

    End Sub

    Sub AddAssay(ByVal boolFromNew As Boolean, ByVal boolFromApply As Boolean)

        'get data from Studies
        Dim intRow As Int16
        Dim intRows As Int16
        Dim intRowA As Int16
        Dim dgv As DataGridView
        Dim dgv1 As DataGridView
        Dim idS As Int64
        Dim idS1
        Dim idT As Int64
        Dim idA As Int64
        Dim boolChkApply As Boolean = False
        Dim strName As String
        Dim str1 As String
        Dim strF As String
        Dim rows() As DataRow

        dgv = Me.dgvSDStudy
        If dgv.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = dgv.CurrentRow.Index
        End If

        idS = dgv("ID_TBLGUWUSTUDIES", intRow).Value


        Dim frm As New frmApplyAssay

        frm.boolFromNew = boolFromNew
        frm.boolFromApply = boolFromApply
        frm.idS = idS

        If boolFromApply Then 'enter assayname
            intRowA = Me.dgvAssays.CurrentRow.Index
            str1 = Me.dgvAssays("CHARASSAYNAME", intRowA).Value
            frm.txtAssayName.Text = str1
            idA = Me.dgvAssays("ID_TBLGUWUASSAY", intRowA).Value
            frm.idA = idA
        Else
            frm.idA = -1
        End If

        Call frm.FormLoad()

        frm.ShowDialog()

        Me.Refresh()

        If frm.boolCancel Then
            frm.Dispose()
            Exit Sub
        Else
            Me.cmdAddAssay.Enabled = False
        End If

        idT = frm.idT

        If frm.chkApplyTemplate.Checked Then
            boolChkApply = True
        Else
            boolChkApply = False
        End If

        strName = frm.txtAssayName.Text

        frm.Dispose()

        Dim tbl As System.Data.Datatable
        Dim maxID As Int64
        Dim Count1 As Short
        Dim Count2 As Short
        Dim boolHit As Boolean
        Dim strM As String
        Dim rowsS() As DataRow
        Dim boolDo As Boolean
        Dim var1, var2

        dgv = Me.dgvSDStudy
        dgv1 = Me.dgvAssays
        intRows = dgv.Rows.Count
        If intRows = 0 Then
            Exit Sub
        End If

        'get data from Studies
        If dgv.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = dgv.CurrentRow.Index
        End If

        tbl = tblGuWuAssay

        idS = dgv("ID_TBLGUWUSTUDIES", intRow).Value
        idS1 = dgv("ID_TBLSTUDIES", intRow).Value

        If boolFromNew And boolChkApply Then 'enter everything
            maxID = GetMaxID("TBLGUWUASSAY", 1, True)
        ElseIf boolFromNew And boolChkApply = False Then 'only add a new asay
            maxID = GetMaxID("TBLGUWUASSAY", 1, True)
        ElseIf boolFromApply Then 'enter only applied fields

        End If

        If boolFromNew And boolChkApply Then 'enter everything
            Dim row As DataRow = tbl.NewRow
            row.BeginEdit()
            row("ID_TBLGUWUASSAY") = maxID
            row("ID_TBLGUWUSTUDIES") = idS
            row("ID_TBLSTUDIES") = idS1
            row("CHARASSAYNAME") = strName

            idA = maxID

            'get template info
            strF = "ID_TBLGUWUASSAY = " & idT
            rows = tbl.Select(strF, Nothing, DataViewRowState.CurrentRows)
            For Count1 = 0 To rows.Length - 1
                For Count2 = 0 To tbl.Columns.Count - 1
                    boolDo = False
                    var1 = rows(0).Item(Count2)
                    str1 = tbl.Columns(Count2).ColumnName
                    Select Case str1
                        Case Is = "ID_TBLGUWUSPECIES"
                            boolDo = True
                        Case Is = "CHARSPECIES"
                            boolDo = True
                        Case Is = "CHARSPECIESSTRAIN"
                            boolDo = True
                        Case Is = "CHARDOSEUNITS"
                            boolDo = True
                        Case Is = "CHARDOSECONCUNITS"
                            boolDo = True
                        Case Is = "CHARTISSUEWTUNITS"
                            boolDo = True
                        Case Is = "CHARPREVPATREQ"
                            boolDo = True
                        Case Is = "CHARSTUDYDESIGNTYPE"
                            boolDo = True
                        Case Is = "BOOLINCLUDE"
                            boolDo = True
                        Case Is = "INTSTATUS"
                            boolDo = True
                        Case Is = "CHARSUBJECT"
                            boolDo = True
                        Case Is = "CHARDESCRIPTION"
                            boolDo = True
                        Case Is = "INTLABEL"
                            boolDo = True
                        Case Is = "CHARLOCATION"
                            boolDo = True
                        Case Is = "BOOLALLDAY"
                            boolDo = True
                        Case Is = "INTEVENTTYPE"
                            boolDo = True
                        Case Is = "CHARRECURRENCEINFO"
                            boolDo = True
                        Case Is = "CHARREMINDERINFO"
                            boolDo = True
                        Case Is = "CHARCONTACTINFO"
                            boolDo = True
                    End Select
                    If boolDo Then
                        row.Item(str1) = var1
                    End If
                Next
            Next
            row.EndEdit()
            tbl.Rows.Add(row)

            Call AddRestOfAssay(idA, idT, idS, idS1)


        ElseIf boolFromNew And boolChkApply = False Then 'only add a new asay
            Dim row As DataRow = tbl.NewRow
            row.BeginEdit()
            row("ID_TBLGUWUASSAY") = maxID
            row("ID_TBLGUWUSTUDIES") = idS
            row("ID_TBLSTUDIES") = idS1
            row("CHARASSAYNAME") = strName
            row.EndEdit()
            tbl.Rows.Add(row)
        ElseIf boolFromApply Then 'enter only applied fields

            'get template info
            strF = "ID_TBLGUWUASSAY = " & idT
            rowsS = tbl.Select(strF, Nothing, DataViewRowState.CurrentRows)

            strF = "ID_TBLGUWUASSAY = " & idA
            rows = tbl.Select(strF, Nothing, DataViewRowState.CurrentRows)

            rows(0).BeginEdit()
            For Count1 = 0 To rows.Length - 1
                For Count2 = 0 To tbl.Columns.Count - 1
                    boolDo = False
                    var1 = rowsS(0).Item(Count2)
                    str1 = tbl.Columns(Count2).ColumnName
                    Select Case str1
                        Case Is = "ID_TBLGUWUSPECIES"
                            boolDo = True
                        Case Is = "CHARSPECIES"
                            boolDo = True
                        Case Is = "CHARSPECIESSTRAIN"
                            boolDo = True
                        Case Is = "CHARDOSEUNITS"
                            boolDo = True
                        Case Is = "CHARDOSECONCUNITS"
                            boolDo = True
                        Case Is = "CHARTISSUEWTUNITS"
                            boolDo = True
                        Case Is = "CHARPREVPATREQ"
                            boolDo = True
                        Case Is = "CHARSTUDYDESIGNTYPE"
                            boolDo = True
                        Case Is = "BOOLINCLUDE"
                            boolDo = True
                        Case Is = "INTSTATUS"
                            boolDo = True
                        Case Is = "CHARSUBJECT"
                            boolDo = True
                        Case Is = "CHARDESCRIPTION"
                            boolDo = True
                        Case Is = "INTLABEL"
                            boolDo = True
                        Case Is = "CHARLOCATION"
                            boolDo = True
                        Case Is = "BOOLALLDAY"
                            boolDo = True
                        Case Is = "INTEVENTTYPE"
                            boolDo = True
                        Case Is = "CHARRECURRENCEINFO"
                            boolDo = True
                        Case Is = "CHARREMINDERINFO"
                            boolDo = True
                        Case Is = "CHARCONTACTINFO"
                            boolDo = True
                    End Select
                    If boolDo Then
                        rows(0).Item(str1) = var1
                    End If
                Next
            Next

            rows(0).EndEdit()

            Call AddRestOfAssay(idA, idT, idS, idS1)

        End If


        'choose new assay
        dgv = Me.dgvAssays
        intRows = dgv.Rows.Count
        For Count1 = 0 To intRows - 1
            str1 = dgv("CHARASSAYNAME", Count1).Value
            If StrComp(str1, strName, CompareMethod.Text) = 0 Then
                intRow = Count1
                Exit For
            End If
        Next

        dgv.CurrentCell = dgv.Rows(intRow).Cells("CHARASSAYNAME")
        dgv.CurrentRow.Selected = True


    End Sub

    Sub AddRestOfAssay(ByVal idA As Int64, ByVal idT As Int64, ByVal idS As Int64, ByVal idS1 As Int64)

        'idA=Destination, idT=Source, idS=id_tblGuWuStudies, idS1=id_tblstudies

        Dim tbl As System.Data.Datatable
        Dim rowD() As DataRow
        Dim rowS() As DataRow
        Dim strFD As String
        Dim strFS As String
        Dim strS As String
        Dim Count1 As Short
        Dim Count2 As Short
        Dim maxID As Int64
        Dim Count3 As Short
        Dim strTable As String
        Dim boolSkip As Boolean


        Dim tblGroups As New System.Data.Datatable
        Dim rowsG() As DataRow
        Dim strFG As String
        Dim strSG As String
        Dim idG As Int64
        Dim id1 As Int64

        Dim col1 As New DataColumn
        col1.AllowDBNull = True
        col1.ColumnName = "IDGROUPS_SOURCE"
        col1.DataType = System.Type.GetType("System.Int64")
        tblGroups.Columns.Add(col1)

        Dim col2 As New DataColumn
        col2.AllowDBNull = True
        col2.ColumnName = "IDGROUPS_DEST"
        col2.DataType = System.Type.GetType("System.Int64")
        tblGroups.Columns.Add(col2)

        Dim col3 As New DataColumn
        col3.AllowDBNull = True
        col3.ColumnName = "IDROUTES_SOURCE"
        col3.DataType = System.Type.GetType("System.Int64")
        tblGroups.Columns.Add(col3)

        Dim col4 As New DataColumn
        col4.AllowDBNull = True
        col4.ColumnName = "IDROUTES_DEST"
        col4.DataType = System.Type.GetType("System.Int64")
        tblGroups.Columns.Add(col4)

        strFD = "ID_TBLGUWUASSAY = " & idA
        strFS = "ID_TBLGUWUASSAY = " & idT

        For Count3 = 1 To 6
            Select Case Count3
                Case 1
                    strTable = "TBLGUWUPKGROUPS"
                    boolSkip = False
                    tbl = tblGuWuPKGroups
                Case 2
                    strTable = "TBLGUWUPKROUTES"
                    boolSkip = False
                    tbl = tblGuWuPKRoutes
                Case 3
                    strTable = "TBLGUWURTTIMEPOINTS"
                    boolSkip = True
                    'tbl = tblGuWuPKGroups
                Case 4
                    strTable = "TBLGUWUPKSUBJECTS"
                    boolSkip = True
                    'tbl = tblGuWuPKGroups
                Case 5 'Compounds
                    strTable = "TBLGUWUASSIGNEDCMPD"
                    boolSkip = False
                    tbl = tblGuWuAssignedCmpd
                Case 6 'Personnel
                    strTable = "TBLGUWUASSAYPERS"
                    boolSkip = False
                    tbl = tblGuWuAssayPERS

            End Select
            If boolSkip Then
                GoTo next1
            End If

            strS = "ID_" & strTable & " ASC"
            rowS = tbl.Select(strFS, strS, DataViewRowState.CurrentRows)
            rowD = tbl.Select(strFD, strS, DataViewRowState.CurrentRows)

            'delete rows and make new ones
            For Count1 = 0 To rowD.Length - 1
                rowD(Count1).Delete()
            Next

            For Count1 = 0 To rowS.Length - 1
                If StrComp(strTable, "TBLGUWUPKGROUPS", CompareMethod.Text) = 0 Then
                    Dim nrowG As DataRow = tblGroups.NewRow
                    nrowG.BeginEdit()
                    maxID = GetMaxID(strTable, 1, True)
                    Dim nrow As DataRow = tbl.NewRow
                    nrow.BeginEdit()
                    For Count2 = 0 To tbl.Columns.Count - 1
                        If Count2 = 0 Then
                            nrow.Item(Count2) = maxID
                            nrowG.Item("IDGROUPS_DEST") = maxID
                            nrowG.Item("IDGROUPS_SOURCE") = rowS(Count1).Item(Count2)
                        Else
                            nrow.Item(Count2) = rowS(Count1).Item(Count2)
                        End If
                    Next

                    nrow.Item("ID_TBLGUWUASSAY") = idA
                    nrow.Item("ID_TBLGUWUSTUDIES") = idS

                    nrowG.EndEdit()
                    tblGroups.Rows.Add(nrowG)

                    nrow.EndEdit()
                    tbl.Rows.Add(nrow)

                ElseIf StrComp(strTable, "TBLGUWUPKROUTES", CompareMethod.Text) = 0 Then
                    maxID = GetMaxID(strTable, 1, True)
                    Dim nrow As DataRow = tbl.NewRow
                    nrow.BeginEdit()

                    id1 = rowS(Count1).Item("ID_TBLGUWUPKGROUPS")
                    strFG = "IDGROUPS_SOURCE = " & id1
                    rowsG = tblGroups.Select(strFG)
                    idG = rowsG(0).Item("IDGROUPS_DEST")

                    For Count2 = 0 To tbl.Columns.Count - 1
                        If Count2 = 0 Then
                            nrow.Item(Count2) = maxID
                        Else
                            nrow.Item(Count2) = rowS(Count1).Item(Count2)
                        End If
                    Next
                    nrow.Item("ID_TBLGUWUPKGROUPS") = idG
                    nrow.Item("ID_TBLGUWUASSAY") = idA
                    nrow.Item("ID_TBLGUWUSTUDIES") = idS

                    nrow.EndEdit()
                    tbl.Rows.Add(nrow)

                ElseIf StrComp(strTable, "TBLGUWUASSIGNEDCMPD", CompareMethod.Text) = 0 Then
                    maxID = GetMaxID(strTable, 1, True)
                    Dim nrow As DataRow = tbl.NewRow
                    nrow.BeginEdit()
                    For Count2 = 0 To tbl.Columns.Count - 1
                        If Count2 = 0 Then
                            nrow.Item(Count2) = maxID
                        Else
                            nrow.Item(Count2) = rowS(Count1).Item(Count2)
                        End If
                    Next
                    nrow.Item("ID_TBLGUWUSTUDIES") = idS
                    nrow.Item("ID_TBLGUWUASSAY") = idA
                    nrow.Item("ID_TBLSTUDIES") = idS1
                    nrow.EndEdit()
                    tbl.Rows.Add(nrow)
                ElseIf StrComp(strTable, "TBLGUWUASSAYPERS", CompareMethod.Text) = 0 Then
                    maxID = GetMaxID(strTable, 1, True)
                    Dim nrow As DataRow = tbl.NewRow
                    nrow.BeginEdit()
                    For Count2 = 0 To tbl.Columns.Count - 1
                        If Count2 = 0 Then
                            nrow.Item(Count2) = maxID
                        Else
                            nrow.Item(Count2) = rowS(Count1).Item(Count2)
                        End If
                    Next
                    nrow.Item("ID_TBLGUWUSTUDIES") = idS
                    nrow.Item("ID_TBLGUWUASSAY") = idA
                    nrow.Item("ID_TBLSTUDIES") = idS1
                    nrow.EndEdit()
                    tbl.Rows.Add(nrow)
                End If

            Next

next1:

        Next

        tblGroups.Dispose()

    End Sub

    Private Sub dgvGroups_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvGroups.SelectionChanged

        If boolFormLoad Then
            Exit Sub
        End If


        boolFromGroupRoute = True

        Call ChangeRoutes()

        Call Filldgv1(Me.dgvGroupDetails, Me.dgvRoutes, "GroupDetails")

        If boolFromGroupSummary Then
        Else
            Call UpdateGroupSummarySelection()
        End If

        Call ChangeTimePoints()

        Call ChangePatients()

        boolFromGroupRoute = False

    End Sub

    Private Sub dgvRoutes_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvRoutes.SelectionChanged

        If boolFormLoad Then
            Exit Sub
        End If

        If boolFromRouteRemove Then
            Exit Sub
        End If

        If boolFromApplyGroup Then
            Exit Sub
        End If

        If boolAssayCancel Then
            Exit Sub
        End If

        '
        boolFromGroupRoute = True

        Call Filldgv1(Me.dgvGroupDetails, Me.dgvRoutes, "GroupDetails")

        If boolFromGroupSummary Then
        Else
            Call UpdateGroupSummarySelection()
        End If

        boolFromGroupRoute = False

        Call ChangeTimePoints()

        Call ChangePatients()


    End Sub

    Sub UpdateGroupSummarySelection()

        If boolFormLoad Then
            Exit Sub
        End If

        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView
        Dim dgv3 As DataGridView
        Dim intRow1 As Short
        Dim intRow2 As Short
        Dim intRow3 As Short
        Dim id1 As Int64
        Dim id2 As Int64
        Dim Count1 As Short
        Dim idA As Int64
        Dim idB As Int64


        dgv1 = Me.dgvGroups
        dgv2 = Me.dgvRoutes
        dgv3 = Me.dgvGroupSummary

        If dgv1.Rows.Count = 0 Then
            Exit Sub
        End If
        If dgv1.CurrentRow Is Nothing Then
            intRow1 = 0
        Else
            intRow1 = dgv1.CurrentRow.Index
        End If
        id1 = dgv1("ID_TBLGUWUPKGROUPS", intRow1).Value

        If dgv2.Rows.Count = 0 Then
            id2 = -1
        ElseIf dgv2.CurrentRow Is Nothing Then
            intRow2 = 0
            id2 = dgv2("ID_TBLGUWUPKROUTES", intRow2).Value
        Else
            intRow2 = dgv2.CurrentRow.Index
            id2 = dgv2("ID_TBLGUWUPKROUTES", intRow2).Value
        End If

        For Count1 = 0 To dgv3.Rows.Count - 1

            idA = dgv3("ID_TBLGUWUPKGROUPS", Count1).Value
            idB = dgv3("ID_TBLGUWUPKROUTES", Count1).Value

            If idA = id1 And idB = id2 Then
                dgv3.CurrentCell = dgv3.Rows(Count1).Cells("ColumnValue")
                dgv3.CurrentRow.Selected = True
                Exit For
            End If

        Next

    End Sub

    Private Sub dgvGroupSummary_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvGroupSummary.SelectionChanged

        Call GroupSummaryChange()


    End Sub

    Sub GroupSummaryChange()

        If boolFormLoad Then
            Exit Sub
        End If

        If boolFromGroupRoute Then
            Exit Sub
        End If

        If boolFromGroupSummary Then
            Exit Sub
        End If

        If boolFromRouteRemove Then
            Exit Sub
        End If

        If boolFromApplyGroup Then
            Exit Sub
        End If

        If boolAssayCancel Then
            Exit Sub
        End If

        boolFromGroupSummary = True

        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView
        Dim dgv3 As DataGridView
        Dim intRow1 As Short
        Dim intRow2 As Short
        Dim intRow3 As Short
        Dim id1 As Int64
        Dim id2 As Int64
        Dim Count1 As Short
        Dim idA As Int64


        dgv1 = Me.dgvGroups
        dgv2 = Me.dgvRoutes
        dgv3 = Me.dgvGroupSummary

        If dgv3.Rows.Count = 0 Then
            GoTo end1
        ElseIf dgv3.CurrentRow Is Nothing Then
            intRow3 = 0
        Else
            intRow3 = dgv3.CurrentRow.Index
        End If

        Call UpdateDGVs("GroupDetails")

        intRow3 = dgv3.CurrentRow.Index
        id1 = dgv3("ID_TBLGUWUPKGROUPS", intRow3).Value
        id2 = dgv3("ID_TBLGUWUPKROUTES", intRow3).Value

        If id1 = -1 Or id2 = -1 Then
            boolFromGroupSummary = False
            Call UpdateGroupSummarySelection()
            GoTo end1
        End If

        For Count1 = 0 To dgv1.Rows.Count - 1
            idA = dgv1("ID_TBLGUWUPKGROUPS", Count1).Value
            If idA = id1 Then
                dgv1.CurrentCell = dgv1.Rows(Count1).Cells("CHARGROUP")
                dgv1.CurrentRow.Selected = True
                Exit For
            End If
        Next

        Call ChangeGroups()

        For Count1 = 0 To dgv2.Rows.Count - 1
            idA = dgv2("ID_TBLGUWUPKROUTES", Count1).Value
            If idA = id2 Then
                dgv2.CurrentCell = dgv2.Rows(Count1).Cells("CHARROUTE")
                dgv2.CurrentRow.Selected = True
                Exit For
            End If
        Next

        Call ChangeRoutes()

        Call ChangeTimePoints()

        Call ChangePatients()

        Call FillPatientsCheck(Me.dgvPatients)

        Call SetSerial()

end1:

        boolFromGroupSummary = False

    End Sub

    Private Sub dgvGroupDetails_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvGroupDetails.CellClick

        If e.ColumnIndex = 2 Then
        Else
            GoTo end1
        End If

        Call ForceCellFormat(Me.dgvGroupDetails)

end1:

    End Sub

    Private Sub dgvGroupDetails_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgvGroupDetails.CellValidating

        If boolFormLoad Then
            Exit Sub
        End If

        If Me.cmdEdit.Enabled Then
            Exit Sub
        End If

        Dim dgv As DataGridView
        Dim intRow As Short
        Dim strM As String
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim boolBool As Boolean = False
        Dim boolID As Boolean = False
        Dim strF As String
        Dim rows() As DataRow
        Dim int1 As Int64
        Dim strDt As String
        Dim strDt1 As String
        Dim strDt2 As String
        Dim intCol As Short

        dgv = Me.dgvGroupDetails
        intCol = e.ColumnIndex
        str1 = dgv.Columns(intCol).Name
        If StrComp(str1, "ColumnValue", CompareMethod.Text) = 0 Then
        Else
            Exit Sub
        End If

        If Len(NZ(e.FormattedValue, "")) = 0 Then
            Exit Sub
        End If

        intRow = dgv.CurrentRow.Index
        str1 = dgv("ColumnName", intRow).Value
        str3 = dgv("ColumnHeader", intRow).Value

        'now transfer data to ColumnValueActual
        If StrComp(Mid(str1, 1, 4), "bool", CompareMethod.Text) = 0 Then
            boolBool = True
        End If

        If StrComp(Mid(str1, 1, 3), "ID_", CompareMethod.Text) = 0 Then
            boolID = True
        End If

        If boolBool Then
            If StrComp(e.FormattedValue, "TRUE", CompareMethod.Text) = 0 Then
                dgv.Item("ColumnValueActual", intRow).Value = -1
            Else
                dgv.Item("ColumnValueActual", intRow).Value = 0
            End If
        ElseIf boolID Then

            'the following is purely for demo purposes
            If StrComp(str1, "ID_TBLGUWUPROJECTS", CompareMethod.Text) = 0 Then


            ElseIf StrComp(str1, "ID_TBLCONFIGREPORTTYPE", CompareMethod.Text) = 0 Then
                If Len(NZ(e.FormattedValue, "")) = 0 Then
                    dgv.Item("ColumnValueActual", intRow).Value = DBNull.Value
                Else
                    strF = "CHARREPORTTYPE = '" & e.FormattedValue & "'"
                    rows = tblConfigReportType.Select(strF)

                    int1 = rows(0).Item("ID_TBLCONFIGREPORTTYPE")
                    If int1 = 0 Then
                        dgv.Item("ColumnValueActual", intRow).Value = DBNull.Value
                    Else
                        dgv.Item("ColumnValueActual", intRow).Value = rows(0).Item("ID_TBLCONFIGREPORTTYPE")
                    End If
                End If

            End If

        ElseIf IsUserIDStuff(str1) Then
            strDt = Mid(str1, 1, 2)
            If StrComp(strDt, "DT", CompareMethod.Text) = 0 Then
                strDt1 = Format(e.FormattedValue, "Long Date")
                strDt2 = Format(e.FormattedValue, "Long Time")
                str1 = strDt1 & " " & strDt2
                dgv.Item("ColumnValueActual", intRow).Value = str1
            Else
                dgv.Item("ColumnValueActual", intRow).Value = e.FormattedValue
            End If
        ElseIf StrComp(str1, "CHARTARGETDOSE", CompareMethod.Text) = 0 Then

            'must be numeric
            If IsNumeric(e.FormattedValue) Then
            Else
                e.Cancel = True
                strM = "Entry must be numeric"
                MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
            End If
            dgv.Item("ColumnValueActual", intRow).Value = e.FormattedValue
        ElseIf StrComp(str1, "CHARTARGETDOSECONC", CompareMethod.Text) = 0 Then

            'must be numeric
            If IsNumeric(e.FormattedValue) Then
            Else
                e.Cancel = True
                strM = "Entry must be numeric"
                MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
            End If
            dgv.Item("ColumnValueActual", intRow).Value = e.FormattedValue
        ElseIf StrComp(str1, "CHARTARGETTISSUEWT", CompareMethod.Text) = 0 Then

            'must be numeric
            If IsNumeric(e.FormattedValue) Then
            Else
                e.Cancel = True
                strM = "Entry must be numeric"
                MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
            End If
            dgv.Item("ColumnValueActual", intRow).Value = e.FormattedValue
        Else
            dgv.Item("ColumnValueActual", intRow).Value = e.FormattedValue
        End If

end1:


    End Sub

    Private Sub mCal2_DateChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DateRangeEventArgs)

    End Sub

    Sub ChangeDateTimePicker()

        'don't change format

        'If Me.rbMonthly.Checked Then
        '    Me.dtp1.CustomFormat = "MMMM"
        '    Me.dtp1.Format = DateTimePickerFormat.Custom
        'ElseIf Me.rbWeekly.Checked Then
        '    Me.dtp1.Format = DateTimePickerFormat.Long
        'ElseIf Me.rbDaily.Checked Then
        '    Me.dtp1.Format = DateTimePickerFormat.Long
        'End If



    End Sub

    Private Sub rbMonthly_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        Call ChangeDateTimePicker()

    End Sub

    Private Sub rbWeekly_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        Call ChangeDateTimePicker()

    End Sub

    Private Sub rbDaily_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        Call ChangeDateTimePicker()

    End Sub

    Sub CreateCalendar()

        Call ConfigCalTable()

        'Dim dgv As DataGridView
        'Dim tbl as System.Data.Datatable
        'Dim boolMon As Boolean = False
        'Dim boolW As Boolean = False
        'Dim boolD As Boolean = False
        'Dim dt1 As Date
        'Dim dt2 As Date
        'Dim dtTP As Date
        'Dim mon As Short
        'Dim var1, var2, var3
        'Dim firstDay As Short
        'Dim lastDay As Short
        'Dim numDays As Short
        'Dim boolPartF As Boolean = True
        'Dim boolPartE As Boolean = True
        'Dim dt1Mon As Date
        'Dim dt2Sun As Date
        'Dim dtT As Date
        'Dim Count1 As Short
        'Dim Count2 As Short
        'Dim int1 As Short
        'Dim intDay As Short
        'Dim str1 As String
        'Dim arr1(2, 6) As Short
        'Dim intMonD As Short
        'Dim intMon As Short
        'Dim intD As Short

        ''1=Row, 2=Column
        ''1=StartA, 2=EndA, 3=StartB, 4=EndB, 5=StartC, 6=EndC

        'arr1(1, 1) = -1
        'arr1(2, 1) = -1

        'If Me.rbMonthly.Checked Then
        '    boolMon = True
        'ElseIf Me.rbWeekly.Checked Then
        '    boolW = True
        'ElseIf Me.rbDaily.Checked Then
        '    boolD = True
        'End If

        'dtTP = Me.dtp1.Value

        'dgv = Me.dgvCal1
        'tbl = Me.tblCal


        'tbl.Clear()

        'If boolMon Then
        '    mon = dtTP.Month
        '    dt1 = DateValue(MonthName(dtTP.Month) & " 01, " & dtTP.Year)
        '    var1 = LastDayInMonth(dtTP.Year, mon)
        '    dt2 = DateValue(MonthName(dtTP.Month) & " " & var1 & ", " & dtTP.Year)

        '    dt1 = dtTP
        '    dtT = DateAdd(DateInterval.Month, 1, dt1)
        '    dt2 = DateAdd(DateInterval.Day, -1, dtT)

        '    numDays = DateDiff(DateInterval.Day, dt1, dt2) + 1

        '    intMonD = Month(dt1)

        '    firstDay = Weekday(dt1) 'eg. Monday
        '    lastDay = Weekday(dt2) 'eg. Monday


        '    If firstDay = 1 Then
        '        boolPartF = False
        '    Else
        '        boolPartF = True
        '    End If

        '    If lastDay = 7 Then
        '        boolPartE = False
        '    Else
        '        boolPartE = True
        '    End If

        '    If boolPartF = False Then
        '        dt1Mon = dt1
        '    Else
        '        'look for first monday
        '        int1 = 0
        '        intDay = 0
        '        Do Until intDay = 1
        '            int1 = int1 - 1
        '            dtT = DateAdd(DateInterval.Day, int1, dt1)
        '            intDay = Weekday(dtT)
        '        Loop
        '        dt1Mon = dtT
        '    End If

        '    If boolPartE = False Then
        '        dt2Sun = dt2
        '    Else
        '        'look for first monday
        '        int1 = 0
        '        intDay = 0
        '        Do Until intDay = 7
        '            int1 = int1 + 1
        '            dtT = DateAdd(DateInterval.Day, int1, dt2)
        '            intDay = Weekday(dtT)
        '        Loop
        '        dt2Sun = dtT
        '    End If

        '    'now add rows to table
        '    int1 = -1
        '    intD = 0
        '    For Count1 = 1 To 6

        '        Dim nrow As DataRow = tbl.NewRow
        '        nrow.BeginEdit()
        '        nrow("CHAR") = CStr(Count1)
        '        For Count2 = 1 To 7
        '            Select Case Count2
        '                Case 1
        '                    str1 = "CHARMON"
        '                Case 2
        '                    str1 = "CHARTUE"
        '                Case 3
        '                    str1 = "CHARWED"
        '                Case 4
        '                    str1 = "CHARTHU"
        '                Case 5
        '                    str1 = "CHARFRI"
        '                Case 6
        '                    str1 = "CHARSAT"
        '                Case 7
        '                    str1 = "CHARSUN"
        '            End Select
        '            int1 = int1 + 1
        '            dtT = DateAdd(DateInterval.Day, int1, dt1Mon)
        '            var1 = Microsoft.VisualBasic.Day(dtT)

        '            '1=Row, 2=Column
        '            '1=StartA, 2=EndA, 3=StartB, 4=EndB, 5=StartC, 6=EndC

        '            If Count1 = 1 And Count2 = 1 Then 'add month
        '                If var1 = 1 Then 'skip A's
        '                    arr1(1, 3) = Count1
        '                    arr1(2, 3) = Count1
        '                Else
        '                    arr1(1, 1) = Count1
        '                    arr1(2, 1) = Count2
        '                End If
        '                var2 = Format(dtT, "MMMM")
        '                var1 = var2 & " " & var1
        '            ElseIf Count1 <> 1 And Count2 <> 1 And var1 = 1 Then 'add month

        '                arr1(1, 2) = Count1
        '                If Count2 = 1 Then
        '                    arr1(1, 2) = Count1 - 1
        '                    arr1(2, 2) = 7
        '                ElseIf Count2 = 7 Then
        '                    arr1(2, 2) = Count2 - 1
        '                Else
        '                    arr1(2, 2) = Count2 - 1
        '                End If

        '                arr1(1, 3) = Count1
        '                arr1(2, 3) = Count2

        '                var2 = Format(dtT, "MMMM")
        '                var1 = var2 & " " & var1
        '            End If
        '            intMon = Month(dtT)

        '            If CInt(var1) = numDays And intMon = intMonD Then
        '                arr1(1, 3) = Count1
        '                arr1(2, 3) = Count1
        '            ElseIf CInt(var1) = numDays And intMon <> intMonD Then
        '                arr1(1, 3) = Count1
        '                arr1(2, 3) = Count1
        '            End If
        '            nrow(str1) = var1
        '        Next

        '        '1=Row, 2=Column
        '        '1=StartA, 2=EndA, 3=StartB, 4=EndB, 5=StartC, 6=EndC

        '        nrow.EndEdit()
        '        tbl.Rows.Add(nrow)
        '    Next
        'End If

        ''format cells
        'For Count1 = 0 To 6
        '    dgv.Rows(Count1).Cells(0).Style.BackColor = Color.Gray
        '    For Count2 = 1 To 7
        '        var1 = dgv(Count2, Count1).Value

        '    Next
        'Next

        'Call FormatCalCells(dt1Mon, dt2Sun, dt1, dt2)

    End Sub

    Sub FormatCalCells(ByVal dt1Mon As Date, ByVal dt2Sun As Date, ByVal dt1 As Date, ByVal dt2 As Date)

        'Dim Count1 As Short
        'Dim Count2 As Short
        'Dim dgv As DataGridView
        'Dim intMonD As Short
        'Dim intMon As Short
        'Dim int1 As Short
        'Dim int2 As Short
        'Dim var1, var2


        'dgv = Me.dgvCal1
        'intMonD = Month(dt1)

        'For Count1 = 0 To 6
        '    dgv.Rows(Count1).Cells(0).Style.BackColor = Color.Gray
        '    For Count2 = 1 To 7
        '        var1 = dgv(Count2, Count1).Value

        '    Next
        'Next


    End Sub


    Function LastDayInMonth(ByVal YearValue As Long, ByVal MonthValue As Long) As Long

        'Rick Rothstein in microsoft.public.vb.general.discussion

        'For the specified year and month,
        'return the last day of that month
        LastDayInMonth = Microsoft.VisualBasic.Day(DateSerial(YearValue, MonthValue + 1, 0))

    End Function

    Private Function IsLastDayOfMonth(ByVal d As Date) As Boolean

        'George Copeland in comp.lang.basic.visual
        IsLastDayOfMonth = CBool(DatePart("m", d) - DatePart("m", (DateAdd("d", 1, d))))

    End Function


    Sub ConfigCalTable()

        'Dim dtbl as System.Data.Datatable

        'dtbl = Me.tblCal

        'If dtbl.Columns.Contains("CHARMON") Then
        '    Exit Sub
        'End If

        'Dim colR As New DataColumn
        'colR.AllowDBNull = True
        'colR.ColumnName = "CHAR"
        'colR.DataType = System.Type.GetType("System.String")
        'colR.Caption = ""
        'dtbl.Columns.Add(colR)

        'Dim col1 As New DataColumn
        'col1.AllowDBNull = True
        'col1.ColumnName = "CHARMON"
        'col1.DataType = System.Type.GetType("System.String")
        'col1.Caption = "Monday"
        'dtbl.Columns.Add(col1)

        'Dim col2 As New DataColumn
        'col2.AllowDBNull = True
        'col2.ColumnName = "CHARTUE"
        'col2.DataType = System.Type.GetType("System.String")
        'col2.Caption = "Tuesday"
        'dtbl.Columns.Add(col2)

        'Dim col3 As New DataColumn
        'col3.AllowDBNull = True
        'col3.ColumnName = "CHARWED"
        'col3.DataType = System.Type.GetType("System.String")
        'col3.Caption = "Wednesday"
        'dtbl.Columns.Add(col3)

        'Dim col4 As New DataColumn
        'col4.AllowDBNull = True
        'col4.ColumnName = "CHARTHU"
        'col4.DataType = System.Type.GetType("System.String")
        'col4.Caption = "Thursday"
        'dtbl.Columns.Add(col4)

        'Dim col5 As New DataColumn
        'col5.AllowDBNull = True
        'col5.ColumnName = "CHARFRI"
        'col5.DataType = System.Type.GetType("System.String")
        'col5.Caption = "Friday"
        'dtbl.Columns.Add(col5)

        'Dim col6 As New DataColumn
        'col6.AllowDBNull = True
        'col6.ColumnName = "CHARSAT"
        'col6.DataType = System.Type.GetType("System.String")
        'col6.Caption = "Saturday"
        'dtbl.Columns.Add(col6)

        'Dim col7 As New DataColumn
        'col7.AllowDBNull = True
        'col7.ColumnName = "CHARSUN"
        'col7.DataType = System.Type.GetType("System.String")
        'col7.Caption = "Sunday"
        'dtbl.Columns.Add(col7)

        'Dim dv as system.data.dataview = New DataView(dtbl)

        'dv.AllowNew = False
        'dv.AllowDelete = False

        'Dim dgv As DataGridView
        'Dim Count1 As Short
        'Dim var1

        'dgv = Me.dgvCal1

        'dgv.DataSource = dv

        'dgv.Columns("CHARMON").HeaderText = "Monday"
        'dgv.Columns("CHARTUE").HeaderText = "Tuesday"
        'dgv.Columns("CHARWED").HeaderText = "Wednesday"
        'dgv.Columns("CHARTHU").HeaderText = "Thursday"
        'dgv.Columns("CHARFRI").HeaderText = "Friday"
        'dgv.Columns("CHARSAT").HeaderText = "Saturday"
        'dgv.Columns("CHARSUN").HeaderText = "Sunday"

        'dgv.RowHeadersWidth = 0

        'var1 = (dgv.Width - 25) / 7
        'dgv.Columns(0).MinimumWidth = 25
        'For Count1 = 1 To 7
        '    dgv.Columns(Count1).MinimumWidth = var1
        '    dgv.Columns(Count1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        'Next

        'dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        'dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopRight

        'dgv.AutoResizeColumns()


    End Sub

    Private Sub dtp1_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        'Call CreateCalendar()

    End Sub

    Private Sub rbDailyS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbDailyS.CheckedChanged
        'Call gbxScheduleChange()
    End Sub

    Private Sub rbWeeklyS_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbWeeklyS.CheckedChanged
        'Call gbxScheduleChange()
    End Sub

    Private Sub rbMonthS_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbMonthS.CheckedChanged
        'Call gbxScheduleChange()
    End Sub

    'Sub gbxScheduleChange()

    '    If Me.rbDailyS.Checked Then
    '        Me.SchedulerControl1.ActiveViewType = DevExpress.XtraScheduler.SchedulerViewType.Day
    '        Me.SchedulerControl1.Views.DayView.Enabled = True
    '        Me.SchedulerControl1.Views.WeekView.Enabled = False
    '        Me.SchedulerControl1.Views.MonthView.Enabled = False
    '    ElseIf Me.rbWeeklyS.Checked Then
    '        Me.SchedulerControl1.ActiveViewType = DevExpress.XtraScheduler.SchedulerViewType.Week
    '        Me.SchedulerControl1.Views.DayView.Enabled = False
    '        Me.SchedulerControl1.Views.WeekView.Enabled = True
    '        Me.SchedulerControl1.Views.MonthView.Enabled = False
    '    ElseIf Me.rbMonthS.Checked Then
    '        Me.SchedulerControl1.ActiveViewType = DevExpress.XtraScheduler.SchedulerViewType.Month
    '        Me.SchedulerControl1.Views.DayView.Enabled = False
    '        Me.SchedulerControl1.Views.WeekView.Enabled = False
    '        Me.SchedulerControl1.Views.MonthView.Enabled = True
    '    End If

    'End Sub

    Private Sub SchedulerControl1_ActiveViewChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        'Call gbxScheduleChange()

    End Sub

    Private Sub dtpSched_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpSched.ValueChanged

        Dim dt1 As Date

        dt1 = Me.dtpSched.Value

        'Me.SchedulerControl1.Start = dt1




    End Sub

    Sub FillSchedFilter()

        Me.cbxSchedFilter.Items.Clear()

        Me.cbxSchedFilter.Items.Add("[NONE]")
        Me.cbxSchedFilter.Items.Add("Date Range")
        Me.cbxSchedFilter.Items.Add("Projects")
        Me.cbxSchedFilter.Items.Add("Studies")
        Me.cbxSchedFilter.Items.Add("PI")
        Me.cbxSchedFilter.Items.Add("Analyst")
        Me.cbxSchedFilter.Items.Add("Compounds")

    End Sub

    Sub LoadScbx()

        Me.cbxSStudies.Left = Me.cbxSProjects.Left
        Me.cbxSStudies.Top = Me.cbxSProjects.Top

        Me.cbxSPersonnel.Left = Me.cbxSProjects.Left
        Me.cbxSPersonnel.Top = Me.cbxSProjects.Top

        Me.cbxSCompound.Left = Me.cbxSProjects.Left
        Me.cbxSCompound.Top = Me.cbxSProjects.Top

        'Me.panWeekRange.Left = Me.cbxSProjects.Left
        'Me.panWeekRange.Top = Me.cbxSProjects.Top

        Me.cbxSProjects.DataSource = tblGuWuProjects
        Me.cbxSProjects.DisplayMember = "CHARPROJECTNAME"

        Me.cbxSStudies.DataSource = tblGuWuStudies
        Me.cbxSStudies.DisplayMember = "CHARSTUDYNAME"

        Me.cbxSPersonnel.DataSource = tblPersonnel
        Me.cbxSPersonnel.DisplayMember = "CHARLASTNAME"

        Me.cbxSCompound.DataSource = tblGuWuCompounds
        Me.cbxSCompound.DisplayMember = "CHARANALYTENAME"

        Me.cbxSchedFilter.SelectedIndex = -1

        Call Me.cbxSVisible()

    End Sub

    Sub cbxSVisible()

        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim boolVis As Boolean

        str1 = NZ(Me.cbxSchedFilter.SelectedItem, "[NONE]")
        boolVis = False
        str2 = ""
        Select Case str1
            Case "[NONE]"
                Me.cbxSProjects.Visible = False
                Me.cbxSStudies.Visible = False
                Me.cbxSPersonnel.Visible = False
                Me.cbxSCompound.Visible = False
                'Me.panWeekRange.Visible = False
                boolVis = False

            Case "Date Range"
                Me.cbxSProjects.Visible = False
                Me.cbxSStudies.Visible = False
                Me.cbxSPersonnel.Visible = False
                Me.cbxSCompound.Visible = False
                'Me.panWeekRange.Visible = True
                boolVis = False

            Case "Projects"
                Me.cbxSProjects.Visible = True
                Me.cbxSStudies.Visible = False
                Me.cbxSPersonnel.Visible = False
                Me.cbxSCompound.Visible = False
                'Me.panWeekRange.Visible = False
                boolVis = True
                str2 = "Choose a Project:"

            Case "Studies"
                Me.cbxSProjects.Visible = False
                Me.cbxSStudies.Visible = True
                Me.cbxSPersonnel.Visible = False
                Me.cbxSCompound.Visible = False
                'Me.panWeekRange.Visible = False
                boolVis = True
                str2 = "Choose a Study:"

            Case "PI"
                Me.cbxSProjects.Visible = False
                Me.cbxSStudies.Visible = False
                Me.cbxSPersonnel.Visible = True
                Me.cbxSCompound.Visible = False
                'Me.panWeekRange.Visible = False
                boolVis = True
                str2 = "Choose a PI:"

            Case "Analyst"
                Me.cbxSProjects.Visible = False
                Me.cbxSStudies.Visible = False
                Me.cbxSPersonnel.Visible = True
                Me.cbxSCompound.Visible = False
                'Me.panWeekRange.Visible = False
                boolVis = True
                str2 = "Choose an Analyst:"

            Case "Compounds"
                Me.cbxSProjects.Visible = False
                Me.cbxSStudies.Visible = False
                Me.cbxSPersonnel.Visible = False
                Me.cbxSCompound.Visible = True
                'Me.panWeekRange.Visible = False
                boolVis = True
                str2 = "Choose a Compound:"

        End Select

        Me.lblcbxS.Text = str2

    End Sub

    'Private Sub SpinEdit1_EditValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SpinEdit1.EditValueChanged

    '    Dim int1 As Short
    '    Dim strM As String

    '    int1 = Me.SpinEdit1.Value
    '    If int1 > 0 And int1 < 7 Then

    '    Else
    '        strM = "Value must be >= 1 and <= 6."
    '        MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
    '        If int1 < 1 Then
    '            Me.SpinEdit1.Value = 1
    '        ElseIf int1 > 6 Then
    '            Me.SpinEdit1.Value = 6
    '        End If
    '    End If

    '    Me.SchedulerControl1.MonthView.WeekCount = Convert.ToInt32(Me.SpinEdit1.EditValue)

    'End Sub

    Private Sub cbxSchedFilter_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSchedFilter.SelectedIndexChanged

        Call cbxSVisible()

    End Sub

    Sub CalendarAct(ByVal strCBX As String)
        Dim tbl As System.Data.Datatable
        Dim strF As String
        Dim rows() As DataRow
        Dim str1 As String
        Dim ID As Int64

        Select Case strCBX
            Case "Projects"
                str1 = Me.cbxSProjects.SelectedText

                strF = "CHARPROJECTNAME = '" & str1 & "'"
                tbl = tblGuWuProjects
                rows = tbl.Select(strF)



            Case "Studies"

            Case "Personnel"

            Case "Compound"

        End Select

    End Sub

    Private Sub cbxSProjects_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSProjects.SelectedIndexChanged



        Call CalendarAct("Projects")



    End Sub

    Private Sub txtLL_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtLL.MouseClick

        Dim var1
        Dim locX, locY

        var1 = NZ(Me.txtLL.Text, "")

        locX = Me.sst1.Left + Me.panWeekRange.Left + Me.txtLL.Left + Me.txtLL.Width
        locY = Me.sst1.Top + Me.panWeekRange.Top

        If IsDate(var1) Then
            Me.mCal1.SelectionStart = var1
            Me.mCal1.SelectionEnd = var1
        Else
            Me.mCal1.SelectionStart = Now
            Me.mCal1.SelectionEnd = Now
        End If
        Me.mCal1.Location = New system.drawing.point(locX, locY)
        boolFromWeekRange = True
        boolLL = True
        boolUL = False
        Me.mCal1.Visible = True

    End Sub

    Private Sub txtUL_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtUL.MouseClick

        Dim var1
        Dim locX, locY

        var1 = NZ(Me.txtUL.Text, "")

        locX = Me.sst1.Left + Me.panWeekRange.Left + Me.txtLL.Left + Me.txtUL.Width
        locY = Me.sst1.Top + Me.panWeekRange.Top

        If IsDate(var1) Then
            Me.mCal1.SelectionStart = var1
            Me.mCal1.SelectionEnd = var1
        Else
            Me.mCal1.SelectionStart = Now
            Me.mCal1.SelectionEnd = Now
        End If
        Me.mCal1.Location = New system.drawing.point(locX, locY)
        boolFromWeekRange = True
        boolLL = False
        boolUL = False
        Me.mCal1.Visible = True
    End Sub


    Private Sub cmdAddGroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddGroup.Click

        Dim idS As Int64
        Dim idA As Int64
        Dim dgv As DataGridView
        Dim intRow As Short
        Dim maxID As Int64
        Dim strName As String
        Dim strM As String
        Dim dv As system.data.dataview
        Dim Count1 As Short
        Dim str1 As String
        Dim boolHit As Boolean


        'ensure no duplicates
        dv = Me.dgvGroups.DataSource
        boolHit = True

        Do Until boolHit = False
            strM = "Please enter a Group Name:"
            strName = ""
            strName = InputBox(strM, "Enter a Group Name:")
            If Len(strName) = 0 Then
                Exit Sub
            End If

            boolHit = False
            For Count1 = 0 To dv.Count - 1
                str1 = dv(Count1).Item("CHARGROUP")
                If StrComp(str1, strName, CompareMethod.Text) = 0 Then
                    boolHit = True
                    Exit For
                End If
            Next
            If boolHit Then
                strM = "The Group Name '" & strName & "' already exists in this Assay."
                strM = strM & ChrW(10) & ChrW(10) & "Please enter a unique name."
                MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
            End If
        Loop

        maxID = GetMaxID("TBLGUWUPKGROUPS", 1, True)

        dgv = Me.dgvAssays
        If dgv.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = dgv.CurrentRow.Index
        End If
        idS = dgv("ID_TBLGUWUSTUDIES", intRow).Value
        idA = dgv("ID_TBLGUWUASSAY", intRow).Value

        Dim tbl As System.Data.Datatable
        tbl = tblGuWuPKGroups
        Dim nrow As DataRow = tbl.NewRow
        nrow.BeginEdit()
        nrow("ID_TBLGUWUPKGROUPS") = maxID
        nrow("ID_TBLGUWUSTUDIES") = idS
        nrow("ID_TBLGUWUASSAY") = idA
        nrow("CHARGROUP") = strName
        nrow.EndEdit()
        tbl.Rows.Add(nrow)

        'select the new group
        dgv = Me.dgvGroups
        intRow = 0
        For Count1 = 0 To dgv.Rows.Count - 1
            str1 = dgv("CHARGROUP", Count1).Value
            If StrComp(str1, strName, CompareMethod.Text) = 0 Then
                intRow = Count1
                Exit For
            End If
        Next

        dgv.CurrentCell = dgv.Rows(intRow).Cells("CHARGROUP")
        dgv.CurrentRow.Selected = True

        Call CreateGroupSummary()

        Call DoAllLabels()

    End Sub

    Private Sub cmdAddRoute_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAddRoute.Click

        Call AddRoute()

        Call DoAllLabels()


    End Sub

    Private Sub cmdRemoveGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdRemoveGroup.Click

        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView
        Dim intRow1 As Short
        Dim intRow2 As Short
        Dim strM As String
        Dim idG As Int64
        Dim strF As String
        Dim Count1 As Short
        Dim tbl1 As System.Data.Datatable
        Dim tbl2 As System.Data.Datatable
        Dim rows1() As DataRow
        Dim rows2() As DataRow
        Dim int1 As Short
        Dim strG As String


        dgv1 = Me.dgvGroups
        dgv2 = Me.dgvRoutes

        tbl1 = tblGuWuPKGroups
        tbl2 = tblGuWuPKRoutes

        If dgv1.Rows.Count = 0 Then
            Exit Sub
        End If

        If dgv1.CurrentRow Is Nothing Then
            strM = "Please select a Group to remove."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If

        intRow1 = dgv1.CurrentRow.Index
        idG = dgv1("ID_TBLGUWUPKGROUPS", intRow1).Value

        strG = dgv1("CHARGROUP", intRow1).Value

        strM = "Are you sure you want to remove " & strG & "?"
        strM = strM & ChrW(10) & ChrW(10) & "This will remove any Routes associated with this Group."
        int1 = MsgBox(strM, MsgBoxStyle.Information + MsgBoxStyle.OkCancel, "Are you sure...")
        If int1 = 1 Then 'continue
        Else
            Exit Sub
        End If

        'first delete routes
        strF = "ID_TBLGUWUPKGROUPS = " & idG
        rows2 = tbl2.Select(strF)
        For Count1 = 0 To rows2.Length - 1
            rows2(Count1).Delete()
        Next

        'now delete group
        rows1 = tbl1.Select(strF)
        rows1(0).Delete()

        dgv1 = Me.dgvGroups
        dgv2 = Me.dgvRoutes

        If dgv1.Rows.Count = 0 Then
        Else
            dgv1.CurrentCell = dgv1.Rows(0).Cells("CHARGROUP")
        End If

        Call CreateGroupSummary()

        Call DoAllLabels()


    End Sub

    Private Sub cmdRemoveRoute_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdRemoveRoute.Click

        Dim dgv2 As DataGridView
        Dim intRow2 As Short
        Dim strM As String
        Dim idR As Int64
        Dim idT As Int64
        Dim strF As String
        Dim Count1 As Short
        Dim tbl2 As System.Data.Datatable
        Dim rows2() As DataRow
        Dim dv As system.data.dataview

        dgv2 = Me.dgvRoutes

        tbl2 = tblGuWuPKRoutes

        If dgv2.Rows.Count = 0 Then
            Exit Sub
        End If

        If dgv2.CurrentRow Is Nothing Then
            strM = "Please select a Route to remove."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If

        intRow2 = dgv2.CurrentRow.Index
        idR = dgv2("ID_TBLGUWUPKROUTES", intRow2).Value

        boolFromRouteRemove = True

        Dim int1 As Short
        'now delete route
        strF = "ID_TBLGUWUPKROUTES = " & idR
        rows2 = tbl2.Select(strF, Nothing, DataViewRowState.CurrentRows)
        int1 = rows2.Length
        rows2(0).Delete()

        'dv = dgv2.DataSource
        'Dim boolT As Boolean
        'boolT = dv.AllowDelete
        'dv.AllowDelete = True
        'For Count1 = 0 To dv.Count - 1
        '    idT = dv(Count1).Item("ID_TBLGUWUPKROUTES")
        '    If idT = idR Then
        '        dv(Count1).Delete()
        '        Exit For
        '    End If
        'Next
        'dv.AllowDelete = boolT

        dgv2 = Me.dgvRoutes

        If dgv2.Rows.Count = 0 Then
        Else
            dgv2.CurrentCell = dgv2.Rows(0).Cells("CHARROUTE")
            dgv2.CurrentRow.Selected = True
        End If

        'Call ChangeGroups()

        Call CreateGroupSummary()

        boolFromRouteRemove = False

        Call DoAllLabels()


    End Sub

    Sub FilllbxRoutes()

        Me.lbxRoute.Items.Clear()

        Me.lbxRoute.Items.Add("IV")
        Me.lbxRoute.Items.Add("PO")
        Me.lbxRoute.Items.Add("SC")
        Me.lbxRoute.Items.Add("IP")

        Dim str1 As String

        str1 = "<- Double-" & ChrW(10) & "click to Add"
        Me.lbllbxRoute.Text = str1

    End Sub

    Private Sub lbxRoute_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbxRoute.DoubleClick

        Call AddRoute()

    End Sub

    Sub AddRoute()

        Dim idS As Int64
        Dim idA As Int64
        Dim idG As Int64
        Dim dgv As DataGridView
        Dim intRow As Short
        Dim maxID As Int64
        Dim strName As String
        Dim strM As String
        Dim dv As system.data.dataview
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim int1 As String
        Dim boolHit As Boolean

        'ensure no duplicates
        dv = Me.dgvRoutes.DataSource
        boolHit = True

        'Do Until boolHit = False
        '    strM = "Please enter a Route Name:"
        '    strName = ""
        '    strName = InputBox(strM, "Enter a Route Name:")
        '    If Len(strName) = 0 Then
        '        Exit Sub
        '    End If

        '    boolHit = False
        '    For Count1 = 0 To dv.Count - 1
        '        str1 = dv(Count1).Item("CHARROUTE")
        '        If StrComp(str1, strName, CompareMethod.Text) = 0 Then
        '            boolHit = True
        '            Exit For
        '        End If
        '    Next
        '    If boolHit Then
        '        strM = "The Route Name '" & strName & "' already exists in this Group."
        '        strM = strM & ChrW(10) & ChrW(10) & "Please enter a unique name."
        '        MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
        '    End If
        'Loop

        str1 = Me.lbxRoute.SelectedItem
        'append appropriate number
        dgv = Me.dgvRoutes
        int1 = 1
        For Count1 = 0 To dgv.Rows.Count - 1
            str2 = dgv("CHARROUTE", Count1).Value
            If InStr(1, str2, str1, CompareMethod.Text) > 0 Then
                int1 = int1 + 1
            End If
        Next
        strName = str1 & "-" & int1

        maxID = GetMaxID("TBLGUWUPKROUTES", 1, True)

        dgv = Me.dgvGroups
        If dgv.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = dgv.CurrentRow.Index
        End If
        idS = dgv("ID_TBLGUWUSTUDIES", intRow).Value
        idA = dgv("ID_TBLGUWUASSAY", intRow).Value
        idG = dgv("ID_TBLGUWUPKGROUPS", intRow).Value

        Dim tbl As System.Data.Datatable
        tbl = tblGuWuPKRoutes
        Dim nrow As DataRow = tbl.NewRow
        nrow.BeginEdit()
        nrow("ID_TBLGUWUPKROUTES") = maxID
        nrow("ID_TBLGUWUPKGROUPS") = idG
        nrow("ID_TBLGUWUSTUDIES") = idS
        nrow("ID_TBLGUWUASSAY") = idA
        nrow("CHARROUTE") = strName
        nrow.EndEdit()
        tbl.Rows.Add(nrow)

        'select the new group
        dgv = Me.dgvRoutes
        intRow = 0
        For Count1 = 0 To dgv.Rows.Count - 1
            str1 = dgv("CHARROUTE", Count1).Value
            If StrComp(str1, strName, CompareMethod.Text) = 0 Then
                intRow = Count1
                Exit For
            End If
        Next

        dgv.CurrentCell = dgv.Rows(intRow).Cells("CHARROUTE")
        dgv.CurrentRow.Selected = True

        Call CreateGroupSummary()

    End Sub

    Private Sub cmdApplyToAllGroups_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdApplyToAllGroups.Click

        Dim frm As New frmApplyDataToGroups

        Call frm.FormLoad()

        frm.ShowDialog()

        Me.Refresh()

        If frm.boolCancel Then
            GoTo end1
        End If

        Dim idG As Int64
        Dim idR As Int64
        Dim id1 As Int64
        Dim id2 As Int64
        Dim arrRG(2, 10) As Int64
        Dim int1 As Short

        'get From
        Dim dgv As DataGridView
        Dim Count1 As Short
        Dim intRows As Short
        dgv = frm.dgvFrom

        For Count1 = 0 To dgv.Rows.Count - 1
            idG = dgv("ID_TBLGUWUPKGROUPS", Count1).Value
            idR = dgv("ID_TBLGUWUPKROUTES", Count1).Value
            If idR = -1 Then
            Else
                Exit For
            End If
        Next

        dgv = frm.dgvTo
        intRows = dgv.Rows.Count
        ReDim arrRG(2, intRows)

        int1 = 0
        For Count1 = 0 To dgv.Rows.Count - 1
            id1 = dgv("ID_TBLGUWUPKGROUPS", Count1).Value
            id2 = dgv("ID_TBLGUWUPKROUTES", Count1).Value
            If id2 = -1 Then
            Else
                int1 = int1 + 1
                arrRG(1, int1) = id1
                arrRG(2, int1) = id2
            End If
        Next

        Call ApplyToAllGroups(idG, idR, arrRG, int1)

end1:

        frm.Dispose()

        Call DoAllLabels()


    End Sub

    Sub ApplyToAllGroups(ByVal idG As Int64, ByVal idR As Int64, ByVal arrRG(,) As Int64, ByVal intRowsD As Short)

        'idA=Destination, idT=Source, idS=id_tblGuWuStudies, idS1=id_tblstudies

        Dim tbl As System.Data.Datatable
        Dim tbl1 As System.Data.Datatable

        Dim rowD() As DataRow
        Dim rowS() As DataRow
        Dim rows1() As DataRow
        Dim strFD As String
        Dim strFS As String
        Dim strS As String
        Dim Count1 As Short
        Dim Count2 As Short
        Dim maxID As Int64
        Dim Count3 As Short
        Dim Count4 As Short
        Dim strTable As String
        Dim boolSkip As Boolean

        Dim rowsG() As DataRow
        Dim strFG As String
        Dim strSG As String
        Dim id1 As Int64
        Dim str1 As String
        Dim str2 As String

        Dim dgv As DataGridView
        Dim intRow As Short
        Dim idA As Int64
        Dim idS As Int64
        Dim strF As String

        Dim boolMaxID As Boolean

        boolFromApplyGroup = True

        'get idA
        dgv = Me.dgvAssays
        intRow = dgv.CurrentRow.Index
        idA = dgv("ID_TBLGUWUASSAY", intRow).Value

        Dim var1

        'get idS
        dgv = Me.dgvSDStudy
        intRow = dgv.CurrentRow.Index
        idS = dgv("ID_TBLGUWUSTUDIES", intRow).Value

        For Count3 = 1 To 3
            Select Case Count3
                Case 1
                    strTable = "TBLGUWUPKGROUPS"
                    boolSkip = True
                    tbl = tblGuWuPKGroups

                Case 2
                    strTable = "TBLGUWUPKROUTES"
                    boolSkip = False
                    tbl = tblGuWuPKRoutes
                    strFS = "ID_TBLGUWUPKROUTES = " & idR ' & "ID_TBLGUWUPKGROUPS AND = " & idG
                    rowS = tbl.Select(strFS)
                Case 3
                    strTable = "TBLGUWURTTIMEPOINTS"
                    boolSkip = False
                    'tbl = tblGuWuPKGroups
                Case 4
                    strTable = "TBLGUWUPKSUBJECTS" 'do not replicate
                    boolSkip = True
                    'tbl = tblGuWuPKGroups

            End Select
            If boolSkip Then
                GoTo next1
            End If

            If Count3 = 2 Then 'TBLGUWUPKROUTES
                'routes already exist, so simply modify existing records
                ''create strFD
                'For Count1 = 1 To intRowsD
                '    If Count1 = 1 Then
                '        strFD = "ID_TBLGUWUPKROUTES = " & arrRG(2, Count1)
                '    Else
                '        strFD = strFD & " OR ID_TBLGUWUPKROUTES = " & arrRG(2, Count1)
                '    End If
                'Next

                For Count1 = 1 To intRowsD
                    strFD = "ID_TBLGUWUPKROUTES = " & arrRG(2, Count1)
                    rowD = tbl.Select(strFD)
                    For Count2 = 0 To rowD.Length - 1
                        rowD(Count2).BeginEdit()
                        For Count4 = 0 To tbl.Columns.Count - 1
                            str1 = tbl.Columns(Count4).ColumnName
                            If StrComp(str1, "ID_TBLGUWUPKROUTES", CompareMethod.Text) = 0 Then
                            ElseIf StrComp(str1, "ID_TBLGUWUPKGROUPS", CompareMethod.Text) = 0 Then
                            ElseIf StrComp(str1, "CHARROUTE", CompareMethod.Text) = 0 Then
                            Else
                                var1 = rowS(0).Item(Count4)
                                rowD(Count2).Item(Count4) = rowS(0).Item(Count4)
                            End If
                        Next
                        rowD(Count2).EndEdit()
                    Next

                Next

            ElseIf Count3 = 3 Then 'TBLGUWURTTIMEPOINTS

                'first remove all existing timepoint data for this assay
                tbl1 = tblGuWuRTTimePoints
                strF = "ID_TBLGUWUASSAY = " & idA & " AND ID_TBLGUWUPKROUTES <> " & idR
                rows1 = tbl1.Select(strF)
                For Count1 = 0 To rows1.Length - 1
                    rows1(Count1).Delete()
                Next

                'now setup rowS
                strF = "ID_TBLGUWUPKROUTES = " & idR
                rowS = tbl1.Select(strF)

                For Count1 = 1 To intRowsD
                    For Count2 = 0 To rowS.Length - 1
                        Dim nRow As DataRow = tbl1.NewRow

                        If Count1 = 1 Then
                            maxID = GetMaxID("TBLGUWURTTIMEPOINTS", rowS.Length, True)
                        Else
                            boolMaxID = True
                            maxID = maxID + 1
                        End If

                        nRow.BeginEdit()
                        nRow("ID_TBLGUWURTTIMEPOINTS") = maxID
                        nRow("ID_TBLGUWUSTUDIES") = idS
                        nRow("ID_TBLSTUDIES") = rowS(Count2).Item("ID_TBLSTUDIES")
                        nRow("ID_TBLGUWUASSAY") = idA
                        nRow("ID_TBLGUWUPKGROUPS") = idG
                        var1 = arrRG(2, Count1)
                        nRow("ID_TBLGUWUPKROUTES") = arrRG(2, Count1)
                        nRow("NUMTIMEPOINT") = rowS(Count2).Item("NUMTIMEPOINT")
                        nRow.EndEdit()

                        tbl1.Rows.Add(nRow)

                    Next
                Next

                If boolMaxID Then
                    boolMaxID = PutMaxID("TBLGUWURTTIMEPOINTS", maxID)
                End If

            End If



next1:

        Next

        boolFromApplyGroup = False

    End Sub

    Private Sub cmdTimePoints_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTimePoints.Click

        Dim idS As Int64
        Dim idS1 As Int64
        Dim idA As Int64
        Dim idG As Int64
        Dim idR As Int64
        Dim id As Int64
        Dim Count1 As Short
        Dim dgv As DataGridView
        Dim dgv1 As DataGridView
        Dim intRow As Short
        Dim strM As String
        Dim strTitle As String

        dgv = Me.dgvRoutes
        dgv1 = Me.dgvSDStudy

        If dgv.RowCount = 0 Then
            strM = "A Group must be configured."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            GoTo end1
        End If

        If dgv.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = dgv.CurrentRow.Index
        End If

        idS = dgv("ID_TBLGUWUSTUDIES", intRow).Value
        idA = dgv("ID_TBLGUWUASSAY", intRow).Value
        idG = dgv("ID_TBLGUWUPKGROUPS", intRow).Value
        idR = dgv("ID_TBLGUWUPKROUTES", intRow).Value

        If dgv1.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = dgv1.CurrentRow.Index
        End If

        idS1 = dgv1("ID_TBLSTUDIES", intRow).Value


        Dim frm As New frmTimePointConfig

        frm.dvTimePoints = Me.dgvGroupTimePoints.DataSource
        frm.idS = idS
        frm.idS1 = idS1
        frm.idA = idA
        frm.idG = idG
        frm.idR = idR

        'create strTitle
        'strTitle = "Time points for Study: " & Me.cbxStudy.Text
        strTitle = "Study: " & Me.cbxStudy.Text
        strTitle = strTitle & ", Assay: " & Me.dgvAssays("CHARASSAYNAME", Me.dgvAssays.CurrentRow.Index).Value
        strTitle = strTitle & ", Group: " & Me.dgvGroups("CHARGROUP", Me.dgvGroups.CurrentRow.Index).Value
        strTitle = strTitle & ", Route: " & Me.dgvRoutes("CHARROUTE", Me.dgvRoutes.CurrentRow.Index).Value

        frm.lblTitle.Text = strTitle

        Call frm.FormLoad()

        frm.ShowDialog()

        Me.Refresh()

        If frm.boolTP Then
            If boolGuWuOracle Then
                'Try
                '    ta_TBLGUWUTPCONFIG.Update(TBLGUWUTPCONFIG)
                'Catch ex As DBConcurrencyException
                '    'ds2005Acc.TBLGUWUTPCONFIG.Merge('ds2005Acc.TBLGUWUTPCONFIG, True)
                'End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_TBLGUWUTPCONFIGAcc.Update(TBLGUWUTPCONFIG)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLGUWUTPCONFIG.Merge('ds2005Acc.TBLGUWUTPCONFIG, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_TBLGUWUTPCONFIGSQLServer.Update(TBLGUWUTPCONFIG)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLGUWUTPCONFIG.Merge('ds2005Acc.TBLGUWUTPCONFIG, True)
                End Try
            End If
        End If

        If frm.boolTPName Then
            If boolGuWuOracle Then
                'Try
                '    ta_TBLGUWUTPNAMESCONFIG.Update(TBLGUWUTPNAMESCONFIG)
                'Catch ex As DBConcurrencyException
                '    'ds2005Acc.TBLGUWUTPNAMESCONFIG.Merge('ds2005Acc.TBLGUWUTPNAMESCONFIG, True)
                'End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_TBLGUWUTPNAMESCONFIGAcc.Update(TBLGUWUTPNAMESCONFIG)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLGUWUTPNAMESCONFIG.Merge('ds2005Acc.TBLGUWUTPNAMESCONFIG, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_TBLGUWUTPNAMESCONFIGSQLServer.Update(TBLGUWUTPNAMESCONFIG)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLGUWUTPNAMESCONFIG.Merge('ds2005Acc.TBLGUWUTPNAMESCONFIG, True)
                End Try
            End If
        End If

        frm.Dispose()

        Call ChangeTimePoints()

        Call DoAllLabels()


end1:

    End Sub

    Sub ChangePatients()

        Dim id As Int64
        Dim intRow As Short
        Dim strF As String
        Dim dv as system.data.dataview
        Dim dgv As DataGridView

        Try
            dv = Me.dgvPatients.DataSource
            dgv = Me.dgvRoutes
            If dgv.Rows.Count = 0 Then
                id = -1
            ElseIf dgv.CurrentRow Is Nothing Then
                id = dgv.Item("ID_TBLGUWUPKROUTES", 0).Value
            Else
                intRow = dgv.CurrentRow.Index
                id = dgv.Item("ID_TBLGUWUPKROUTES", intRow).Value
            End If
            strF = "ID_TBLGUWUPKROUTES = " & id
            dv.RowFilter = strF

            Try
                'now fill cmpds
                Call FillPatient(dv)
            Catch ex As Exception

            End Try

            Call DoLabel(Me.dgvPatients, Me.lblPatients, "Patients")


        Catch ex As Exception

        End Try


    End Sub

    Sub ChangeTimePoints()

        Dim id As Int64
        Dim intRow As Short
        Dim strF As String
        Dim dv as system.data.dataview
        Dim dgv As DataGridView

        Try
            dv = Me.dgvGroupTimePoints.DataSource
            dgv = Me.dgvRoutes
            If dgv.Rows.Count = 0 Then
                id = -1
            ElseIf dgv.CurrentRow Is Nothing Then
                id = dgv.Item("ID_TBLGUWUPKROUTES", 0).Value
            Else
                intRow = dgv.CurrentRow.Index
                id = dgv.Item("ID_TBLGUWUPKROUTES", intRow).Value
            End If
            strF = "ID_TBLGUWUPKROUTES = " & id
            dv.RowFilter = strF

            Call DoLabel(Me.dgvGroupTimePoints, Me.lblTimePoints, "Time Points")


        Catch ex As Exception

        End Try


    End Sub

    Private Sub cmdAddCompound_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddCompound.Click

        Dim frm As New frmConfigSDCmpds
        Dim Count1 As Short
        Dim dgvS As DataGridView
        Dim dgvD As DataGridView
        Dim dgv As DataGridView
        Dim intRow As Int16
        Dim idS As Int64
        Dim idS1 As Int64
        Dim idA As Int64

        dgvS = Me.dgvCmpd
        dgvD = frm.dgvCmpd

        dgvD.DataSource = dgvS.DataSource
        For Count1 = 0 To dgvS.Columns.Count - 1
            dgvD.Columns(Count1).Visible = dgvS.Columns(Count1).Visible
            dgvD.Columns(Count1).HeaderText = dgvS.Columns(Count1).HeaderText
            dgvD.Columns(Count1).ReadOnly = dgvS.Columns(Count1).ReadOnly
            dgvD.Columns(Count1).DisplayIndex = dgvS.Columns(Count1).DisplayIndex
        Next

        dgv = Me.dgvAssays
        If dgv.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = dgv.CurrentRow.Index
        End If
        idS = dgv("ID_TBLGUWUSTUDIES", intRow).Value
        frm.idS = idS
        idS1 = dgv("ID_TBLSTUDIES", intRow).Value
        frm.idS1 = idS1
        idA = dgv("ID_TBLGUWUASSAY", intRow).Value
        frm.idA = idA

        Call frm.FormLoad()

        frm.ShowDialog()

        If frm.boolCancel Then
            Call ChangeCmpds()
        End If

        frm.Dispose()

        Call DoAllLabels()


    End Sub

    Private Sub dgvCmpd_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvCmpd.SelectionChanged

        Call ChangeCmpdLots()

    End Sub

    Private Sub cmdGetLotNum_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGetLotNum.Click


        Dim Count1 As Short
        Dim dgvS As DataGridView
        Dim dgvD As DataGridView
        Dim dgv As DataGridView
        Dim dgvC As DataGridView
        Dim intRow As Int16
        Dim idS As Int64
        Dim idS1 As Int64
        Dim idA As Int64
        Dim idC As Int64
        Dim idC1 As Int64
        Dim strC As String
        Dim str1 As String
        Dim str2 As String

        dgvC = Me.dgvCmpd

        If dgvC.RowCount = 0 Then
            strC = "Must have a compound configured first."
            MsgBox(strC, MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If

        Dim frm As New frmConfigCmpdLot

        dgvS = Me.dgvLotNum
        dgvD = frm.dgvLot

        dgvD.DataSource = dgvS.DataSource
        For Count1 = 0 To dgvS.Columns.Count - 1
            dgvD.Columns(Count1).Visible = dgvS.Columns(Count1).Visible
            dgvD.Columns(Count1).HeaderText = dgvS.Columns(Count1).HeaderText
            dgvD.Columns(Count1).ReadOnly = dgvS.Columns(Count1).ReadOnly
        Next

        If dgvC.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = dgvC.CurrentRow.Index
        End If
        idS = dgvC("ID_TBLGUWUSTUDIES", intRow).Value
        frm.idS = idS
        idS1 = dgvC("ID_TBLSTUDIES", intRow).Value
        frm.idS1 = idS1
        idA = dgvC("ID_TBLGUWUASSAY", intRow).Value
        frm.idA = idA
        idC = dgvC("ID_TBLGUWUASSIGNEDCMPD", intRow).Value
        frm.idC = idC
        idC1 = dgvC("ID_TBLGUWUCOMPOUNDS", intRow).Value
        frm.idC1 = idC1

        'make strC
        str1 = dgvC("CHARCOMPOUND", intRow).Value
        str2 = dgvC("CHARCOMPANYID", intRow).Value
        strC = str1 & " - " & str2
        frm.txtCmpd.Text = strC

        Call frm.FormLoad()

        frm.ShowDialog()

        If frm.boolCancel Then
            Call ChangeCmpdLots()

        End If

        frm.Dispose()

    End Sub

    Private Sub cmdConfigPI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdConfigPI.Click

        Dim Count1 As Short
        Dim dgvS As DataGridView
        Dim dgvD As DataGridView
        Dim dgv As DataGridView
        Dim dgvC As DataGridView
        Dim intRow As Int16
        Dim idS As Int64
        Dim idS1 As Int64
        Dim idA As Int64
        Dim idC As Int64
        Dim idC1 As Int64
        Dim strC As String
        Dim str1 As String
        Dim str2 As String

        dgvC = Me.dgvAssays

        If dgvC.RowCount = 0 Then
            strC = "Must have an assay configured first."
            MsgBox(strC, MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If

        Dim frm As New frmConfigAssayPers

        frm.dgvPI.DataSource = Me.dgvPI.DataSource
        frm.dgvAnal.DataSource = Me.dgvAnalyst.DataSource

        dgvS = Me.dgvPI
        dgvD = frm.dgvPI

        dgvD.DataSource = dgvS.DataSource
        For Count1 = 0 To dgvS.Columns.Count - 1
            dgvD.Columns(Count1).Visible = dgvS.Columns(Count1).Visible
            dgvD.Columns(Count1).HeaderText = dgvS.Columns(Count1).HeaderText
            dgvD.Columns(Count1).ReadOnly = dgvS.Columns(Count1).ReadOnly
        Next


        dgvS = Me.dgvAnalyst
        dgvD = frm.dgvAnal

        dgvD.DataSource = dgvS.DataSource
        For Count1 = 0 To dgvS.Columns.Count - 1
            dgvD.Columns(Count1).Visible = dgvS.Columns(Count1).Visible
            dgvD.Columns(Count1).HeaderText = dgvS.Columns(Count1).HeaderText
            dgvD.Columns(Count1).ReadOnly = dgvS.Columns(Count1).ReadOnly
        Next


        If dgvC.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = dgvC.CurrentRow.Index
        End If
        idS = dgvC("ID_TBLGUWUSTUDIES", intRow).Value
        frm.idS = idS
        idS1 = dgvC("ID_TBLSTUDIES", intRow).Value
        frm.idS1 = idS1
        idA = dgvC("ID_TBLGUWUASSAY", intRow).Value
        frm.idA = idA

        Call frm.FormLoad()

        frm.ShowDialog()

        Me.Refresh()

        frm.Dispose()

        Call DoAllLabels()


    End Sub

    Sub DoAllLabels()

        Call DoLabel(Me.dgvGroupTimePoints, Me.lblTimePoints, "Time" & ChrW(10) & "Points")
        Call DoLabel(Me.dgvPI, Me.lblPI, "PI's")
        Call DoLabel(Me.dgvAnalyst, Me.lblAnalyst, "Analysts")
        Call DoLabel(Me.dgvCmpd, Me.lblCmpd, "Compounds")
        Call DoLabel(Me.dgvRoutes, Me.lbldgvRoutes, "Routes")
        Call DoLabel(Me.dgvGroups, Me.lbldgvGroups, "Groups")
        Call DoLabel(Me.dgvAssays, Me.lbldgvAssays, "Assays")
        Call DoLabel(Me.dgvGroupTimePoints, Me.lblTimePoints, "Time Points")

    End Sub

    Private Sub cmdPatients_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPatients.Click

        Dim idS As Int64
        Dim idS1 As Int64
        Dim idA As Int64
        Dim idG As Int64
        Dim idR As Int64
        Dim id As Int64
        Dim Count1 As Short
        Dim dgv As DataGridView
        Dim dgv1 As DataGridView
        Dim intRow As Short
        Dim strM As String
        Dim strTitle As String

        dgv = Me.dgvRoutes
        dgv1 = Me.dgvSDStudy

        If dgv.RowCount = 0 Then
            strM = "A Group must be configured."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            GoTo end1
        End If

        If dgv.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = dgv.CurrentRow.Index
        End If

        idS = dgv("ID_TBLGUWUSTUDIES", intRow).Value
        idA = dgv("ID_TBLGUWUASSAY", intRow).Value
        idG = dgv("ID_TBLGUWUPKGROUPS", intRow).Value
        idR = dgv("ID_TBLGUWUPKROUTES", intRow).Value

        If dgv1.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = dgv1.CurrentRow.Index
        End If

        idS1 = dgv1("ID_TBLSTUDIES", intRow).Value

        Dim frm As New frmConfigPatients
        Dim dgvS As DataGridView
        Dim dgvD As DataGridView
        Dim prop As [Property]

        frm.idS = idS
        frm.idS1 = idS1
        frm.idA = idA
        frm.boolFormLoad = True

        dgvS = Me.dgvGroupSummary
        dgvD = frm.dgvGroupSummary

        dgvD.DataSource = dgvS.DataSource
        For Count1 = 0 To dgvS.Columns.Count - 1
            dgvD.Columns(Count1).ReadOnly = dgvS.Columns(Count1).ReadOnly
            dgvD.Columns(Count1).Visible = dgvS.Columns(Count1).Visible
            dgvD.Columns(Count1).HeaderText = dgvS.Columns(Count1).HeaderText
            dgvD.Columns(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvD.Columns(Count1).DefaultCellStyle.Alignment = dgvS.Columns(Count1).DefaultCellStyle.Alignment
            dgvD.RowHeadersWidth = dgvS.RowHeadersWidth
        Next

        dgvS = Me.dgvGroupTimePoints
        dgvD = frm.dgvGroupTimePoints

        dgvD.DataSource = dgvS.DataSource
        For Count1 = 0 To dgvS.Columns.Count - 1
            dgvD.Columns(Count1).ReadOnly = dgvS.Columns(Count1).ReadOnly
            dgvD.Columns(Count1).Visible = dgvS.Columns(Count1).Visible
            dgvD.Columns(Count1).HeaderText = dgvS.Columns(Count1).HeaderText
            dgvD.Columns(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvD.Columns(Count1).DefaultCellStyle.Alignment = dgvS.Columns(Count1).DefaultCellStyle.Alignment
            dgvD.RowHeadersWidth = dgvS.RowHeadersWidth
        Next

        dgvS = Me.dgvPatients
        dgvD = frm.dgvPatients

        Dim dv as system.data.dataview
        'dv = dgvD.DataSource
        'dv.AllowEdit = True

        dgvD.DataSource = dgvS.DataSource
        For Count1 = 0 To dgvS.Columns.Count - 1
            dgvD.Columns(Count1).ReadOnly = dgvS.Columns(Count1).ReadOnly
            dgvD.Columns(Count1).Visible = dgvS.Columns(Count1).Visible
            dgvD.Columns(Count1).HeaderText = dgvS.Columns(Count1).HeaderText
            dgvD.Columns(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvD.Columns(Count1).DefaultCellStyle.Alignment = dgvS.Columns(Count1).DefaultCellStyle.Alignment
            dgvD.RowHeadersWidth = dgvS.RowHeadersWidth
        Next

        Call frm.FormLoad()

        frm.ShowDialog()

        dv = dgvS.DataSource
        dv.AllowEdit = False

        Me.Refresh()

        frm.Dispose()

        Call GroupSummaryChange()

end1:

    End Sub

    Sub SetSerial()

        Dim dgv As DataGridView
        Dim boolS As Boolean
        Dim intRow As Short
        Dim intS As Short

        dgv = Me.dgvPatients
        boolS = True

        If dgv.RowCount = 0 Then
            boolS = True
        Else
            If dgv.CurrentRow Is Nothing Then
                intRow = 0
            Else
                intRow = dgv.CurrentRow.Index
            End If
            intS = -1
            intS = dgv("BOOLSERIALBLEED", intRow).Value
            If intS = -1 Then
                boolS = True
            Else
                boolS = False
            End If
        End If

        If boolS Then
            Me.rbSerial.Checked = True
            dgv.Columns("NUMTIMEPOINT").Visible = False
        Else
            Me.rbSerialNon.Checked = True
            dgv.Columns("NUMTIMEPOINT").Visible = True
        End If


    End Sub

    Private Sub cbxSStudies_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSStudies.SelectedIndexChanged

        Call CalendarAct("Studies")


    End Sub

    Private Sub cbxSPersonnel_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSPersonnel.SelectedIndexChanged

        Call CalendarAct("Personnel")

    End Sub

    Private Sub cbxSCompound_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSCompound.SelectedIndexChanged

        Call CalendarAct("Compound")

    End Sub

    Private Sub mCal1_DateChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles mCal1.DateChanged

    End Sub


End Class