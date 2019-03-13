Option Compare Text

Imports System.Collections.Generic

Public Class frmAdministration
    'NOTES:
    '1. Remove records from tblConfiguration
    '   1	Global	Directory Paths	Report Templates	C:\Labintegrity\StudyDoc\ReportTemplates\	1	0	-1	0	0	5/1/2006
    '   5	Global	Directory Paths	GuWu Statements	C:\Labintegrity\StudyDoc\ReportStatements\ReportStatementsGuWu01.doc	3	0	-1	-1	0	5/1/2006

    Public boolFormLoad As Boolean
    Public boolAddUser As Boolean
    Public boolAddUserAccount As Boolean
    Public incrStart As Int64
    Public incr1 As Int64
    Public boolAddAcct As Boolean
    Public boolUserAcctGo As Boolean
    Public boolCVNeeded As Boolean
    Public boolAddRow As Boolean
    Public boolCancelAddresses As Boolean
    Public boolAddAddresses As Boolean
    Public boolStopItemCheck As Boolean
    Public boolStopPswdCheck As Boolean
    Public boolSave As Boolean
    Public boolDirty As Boolean = False
    Public boolHold As Boolean = False
    Public frmName As String
    Private rowUA As Short = 0
    Private rowUID As Short = 0
    Public ctPswd As Short = 0
    Public arrPswd(1, 1)
    Public boolPermLoad As Boolean = True

    Private boolESigTemp As Boolean = False
    Private boolAuditTrailTemp As Boolean = False

    Public LDAPUserID As String = ""
    Public LDAPPswd As String = ""

    Public boolFromWatsonButton As Boolean = False

    ' Nick Addition
    Sub SetToEditMode()

        Me.cmdEdit.Enabled = False
        Me.cmdEdit.BackColor = System.Drawing.Color.Gray
        Me.cmdSave.Enabled = True
        Me.cmdSave.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCancel.Enabled = True
        Me.cmdCancel.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdExit.Enabled = False
        Me.cmdExit.BackColor = System.Drawing.Color.Gray

    End Sub

    Sub SetToNonEditMode()

        Me.cmdEdit.Enabled = True
        Me.cmdEdit.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdSave.Enabled = False
        Me.cmdSave.BackColor = System.Drawing.Color.Gray
        Me.cmdCancel.Enabled = False
        Me.cmdCancel.BackColor = System.Drawing.Color.Gray
        Me.cmdExit.Enabled = True
        Me.cmdExit.BackColor = System.Drawing.Color.Gainsboro

    End Sub
    Sub FillFCAdmin()

        Try

            Dim dgv As DataGridView
            Dim dv As System.Data.DataView
            Dim strF As String
            Dim strS As String
            Dim Count1 As Short
            Dim Count2 As Short
            Dim str1 As String
            Dim str2 As String
            Dim int1 As Short

            dgv = Me.dgvFC

            strF = "BOOLCUSTOM = -1"
            strS = "ID_TBLFIELDCODES ASC"

            dv = New DataView(tblFieldCodes, strF, strS, DataViewRowState.CurrentRows)

            dv.AllowDelete = False
            dv.AllowEdit = False
            dv.AllowNew = False

            dgv.DataSource = dv

            For Count1 = 0 To dgv.Columns.Count - 1
                dgv.Columns(Count1).Visible = False
            Next

            str1 = "Field Code"
            str2 = "CHARFIELDCODE"
            dgv.Columns(str2).HeaderText = str1
            dgv.Columns(str2).Visible = True
            'dgv.Columns(str2).ReadOnly = False

            str1 = "Description (optional)"
            str2 = "CHARDESCRIPTION"
            dgv.Columns(str2).HeaderText = str1
            dgv.Columns(str2).Visible = True
            'dgv.Columns(str2).ReadOnly = False

            str1 = "Example (optional)"
            str2 = "CHAREXAMPLE"
            dgv.Columns(str2).HeaderText = str1
            dgv.Columns(str2).Visible = True
            'dgv.Columns(str2).ReadOnly = False
            'change order

            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            dgv.AutoResizeColumns()

            dgv.RowHeadersWidth = 25

        Catch ex As Exception

        End Try


    End Sub

    Sub FinalFrozen()

        Dim dgv As DataGridView

        Try
            dgv = Me.dgvUserAttributes
            dgv.Columns("charUserid").Frozen = True
        Catch ex As Exception

        End Try

        Try
            dgv = Me.dgvUsers
            dgv.Columns("charFirstName").Frozen = True
        Catch ex As Exception

        End Try

    End Sub

    Sub SetPanels()

        Dim t, l, w, h
        Dim pan As Panel
        Dim Count1 As Short
        Dim var1

        Dim fw, fh

        'Me.WindowState = FormWindowState.Maximized
        'set form to avail screen
        'w = My.Computer.Screen.WorkingArea.Width
        'h = My.Computer.Screen.WorkingArea.Height

        'Me.Top = 0
        'Me.Left = 0
        'Me.Width = w
        'Me.Height = h

        fw = Me.Width
        fh = Me.Height

        Dim bw As Int16 = (Me.Width - Me.ClientSize.Width) / 2 'form border width
        Dim tbh As Int16 = Me.Height - Me.ClientSize.Height - 2 * bw 'titlebar height

        t = Me.lbxTab1.Top '39
        l = Me.lbxTab1.Left + Me.lbxTab1.Width + 5 ' 176
        w = 795
        'w = fw - (Me.lbxTab1.Left + Me.lbxTab1.Width) - 10 - Me.panSG.Width - 40
        w = fw - (Me.lbxTab1.Left + Me.lbxTab1.Width) - 20 - bw



        h = Me.lbxTab1.Height ' 625
        Me.lbxTab1.Anchor = AnchorStyles.Bottom + AnchorStyles.Top + AnchorStyles.Left

        For Count1 = 1 To 9
            Select Case Count1
                Case 1
                    pan = Me.pan1
                Case 2
                    pan = Me.pan2
                Case 3
                    pan = Me.pan3
                Case 4
                    pan = Me.pan4
                Case 5
                    pan = Me.pan5
                Case 6
                    pan = Me.pan6
                Case 7
                    pan = Me.pan7b
                Case 8
                    pan = Me.panFC
                Case 9
                    pan = Me.pan8
            End Select

            pan.Left = l
            pan.Top = t
            pan.Width = w
            pan.Height = h

            'set anchor
            pan.Anchor = AnchorStyles.Right + AnchorStyles.Bottom + AnchorStyles.Top + AnchorStyles.Left

            'var1 = pan.Visible
            'var1 = var1

        Next

        'Me.lblOpen.Left = l
        'Me.lblOpen.Top = t
        'Me.lblOpen.Width = w
        'Me.lblOpen.Height = h

        'Me.lblOpen.Visible = True

        'set lblRestrictions
        'Me.lblRestricted.Left = (pan.Left + pan.Width) - Me.lblRestricted.Width
        'Me.lblRestricted.BackColor = Color.White

        'set height of panSG
        Dim a, b, c, d, e

        a = Me.pan1.Top + Me.pan1.Height
        b = Me.cmdEdit.Top + Me.cmdEdit.Height
        c = Me.pan1.Left + Me.pan1.Width


        'do some other stuff
        Call SetGlobalBrowse()


    End Sub

    Sub SetGlobalBrowse()

        Me.cmdBrowseGlobal.Left = Me.lblGlobalValues.Left + Me.lblGlobalValues.Width + Me.panGP.Left
        Me.cmdBrowseGlobal.Top = Me.lblGlobalParameters.Top + Me.lblGlobalParameters.Height + 2

    End Sub

    Sub Setlv()

        Me.lvPermissionsAdmin.Left = Me.lvPermissions.Left
        Me.lvPermissionsAdmin.Top = Me.lvPermissions.Top
        Me.lvPermissionsAdmin.Width = Me.lvPermissions.Width
        Me.lvPermissionsAdmin.Height = Me.lvPermissions.Height

        Me.lvPermissionsReportTemplate.Left = Me.lvPermissions.Left
        Me.lvPermissionsReportTemplate.Top = Me.lvPermissions.Top
        Me.lvPermissionsReportTemplate.Width = Me.lvPermissions.Width
        Me.lvPermissionsReportTemplate.Height = Me.lvPermissions.Height

        Me.lvPermissionsFinalReport.Left = Me.lvPermissions.Left
        Me.lvPermissionsFinalReport.Top = Me.lvPermissions.Top
        Me.lvPermissionsFinalReport.Width = Me.lvPermissions.Width
        Me.lvPermissionsFinalReport.Height = Me.lvPermissions.Height

        Me.lvPermissions.BringToFront()
        Me.lvPermissionsAdmin.BringToFront()
        Me.lvPermissionsReportTemplate.BringToFront()
        Me.lvPermissionsFinalReport.BringToFront()

    End Sub

    Sub FormLoad()

        Dim str1 As String
        Dim str2 As String
        Dim ct1 As Short
        Dim Count1 As Short
        'Dim frm As New frmHome_01
        Dim var1, var2, var3
        Dim tbl As System.Data.DataTable
        Dim row() As DataRow

        boolFormLoad = True
        boolAddUser = False
        boolAddUserAccount = False
        incrStart = 999999
        incr1 = incrStart
        boolAddAcct = False
        boolUserAcctGo = True
        boolCVNeeded = True
        boolAddRow = False
        boolCancelAddresses = False
        boolAddAddresses = False
        boolStopItemCheck = False
        boolStopPswdCheck = False
        boolSave = False

        Call SetPanels()
        Call Setlv()

        str1 = "Select Tab Pages in the Template Attributes table below whose contents are to be included in the template."
        Me.lblTExpl.Text = str1

        str1 = "User ID Permissions" & ChrW(10) & "(Check = Allow Editing in that Tab/Window)"
        str1 = "User ID Permissions (Check = Allow Editing in that Tab/Window)"
        Me.lblPermissions.Text = str1

        str1 = "Password integrity policy enforces that a password content meets at least three of the four following criteria:"
        str1 = str1 & Chr(10) & Chr(10)
        str1 = str1 & "      An upper case letter" & Chr(10)
        str1 = str1 & "      A lower case letter" & Chr(10)
        str1 = str1 & "      A digit" & Chr(10)
        str1 = str1 & "      A non-alphanumeric character"

        str1 = str1 & ChrW(10) & ChrW(10) & "Note that all Password Settings are ignored for a StudyDoc account if that account uses Windows Authentication (see the User Accounts page)."

        Me.lblIntegrity.Text = str1

        tbl = tblTab1
        str1 = "intForm = 2"
        str2 = "intOrder ASC"

        row = tbl.Select(str1, str2)

        'fill lbxTab1
        ct1 = row.Length
        Me.lbxTab1.Items.Clear()
        For Count1 = 0 To ct1 - 1
            str1 = row(Count1).Item("charItem")
            Me.lbxTab1.Items.Add(str1)
        Next


        Dim tp, lf, ht, wd
        tp = Me.Top
        lf = Me.Left
        ht = Me.Height
        wd = Me.Width


        'set active status
        Me.rbShowActiveUserIDs.Checked = True
        Me.rbShowActiveUsers.Checked = True
        Me.rbShowActiveTemplates.Checked = True
        Me.rbShowActiveAddresses.Checked = True

        'Call FilllbxSymbol(Me.lbxSymbol, Me.lblSymbol1)


        'add unbound column to tblConfiguration
        If tblConfiguration.Columns.Contains("Example") Then
        Else
            Dim col1 As New DataColumn
            col1.ColumnName = "Example"
            col1.Caption = "Example"
            col1.DataType = System.Type.GetType("System.String")
            tblConfiguration.Columns.Add(col1)
        End If

        boolStopPswdCheck = True

        Call UserAccountInitialize()
        Call UserAccountConfigure()

        boolStopPswdCheck = False

        Call TemplatesInitialize()

        Call CorporateAddressesInitialize()

        Call DropdownboxInitialize()

        Call GlobalInitialize()
        Call GlobalConfigure()

        Call HooksInitialize()

        Call LoadcbxModules()

        Call SelectcbxModules()

        str1 = Me.cbxModules.SelectedItem
        Select Case str1
            Case "StudyDoc Administration"
                Call PasswordInitialize()
                Call GlobalConfigure()

            Case "Report Writer"
                Call GlobalInitialize()
                Call GlobalConfigure()
        End Select


        'lock up form if appropriate
        Dim rows() As DataRow
        Dim bool As Boolean
        Dim strF As String
        Dim int1 As Short

        strF = "ID_TBLPERMISSIONS = " & id_tblPermissions
        int1 = tblPermissions.Rows.Count 'for testing
        rows = tblPermissions.Select(strF)
        bool = False
        If rows.Length = 0 Then
            bool = False
        Else
            bool = True
        End If
        Me.cmdEdit.Enabled = bool

        Me.lvPermissions.View = System.Windows.Forms.View.List
        Me.lvPermissions.CheckBoxes = True

        Me.lvPermissionsAdmin.View = System.Windows.Forms.View.List
        Me.lvPermissionsAdmin.CheckBoxes = True

        Me.lvPermissionsReportTemplate.View = System.Windows.Forms.View.List
        Me.lvPermissionsReportTemplate.CheckBoxes = True

        Me.lvPermissionsFinalReport.View = System.Windows.Forms.View.List
        Me.lvPermissionsFinalReport.CheckBoxes = True


        'pesky
        Call UserAccountConfigure()
        Call GlobalConfigure()

        Me.lbxTab1.Select()

        'select first row of lbx
        Dim boolF As Boolean = boolFormLoad
        boolFormLoad = False
        Me.lbxTab1.SelectedIndex = 0
        Call lbxTab1Change()

        Call FillFCAdmin()

        boolFormLoad = boolF

        Call LockAll(True)

        'pesky
        Call TemplatesAttributesConfigure()


        Call ShowRFC()
        Call ShowMOS()

        'pesky
        Call ConfigureUserAccountAttributes(False)
        Call GlobalConfigure()

        Call SetPanels()

        Call PositionProgress()
        'SendKeys.Send("%")
    End Sub

    Sub PositionProgress()

        Me.lblPermissions.Width = Me.pb1.Width

        Dim a, b, c, d

        a = Me.pan1.Top + Me.pan1.Height
        b = Me.pan1.Left + Me.pan1.Width

        Me.lblProgress.Width = Me.pan1.Width ' b
        Me.lblProgress.Left = Me.pan1.Left ' (b / 2) - (Me.lblProgress.Width / 2)
        Me.lblProgress.Top = (a / 2) - (Me.lblProgress.Height / 2)

        a = Me.pan1.Top
        Me.lblProgress.Top = a

        Me.lblProgress.Height = Me.pan1.Height - Me.pb1.Height - 2

        Me.pb1.Width = Me.lblProgress.Width
        Me.pb1.Left = Me.lblProgress.Left
        Me.pb1.Top = Me.lblProgress.Top + Me.lblProgress.Height + 2



        Me.lblProgress.BringToFront()
        Me.pb1.BringToFront()

    End Sub

    Sub SelectcbxModules()

        Dim Count1 As Short
        Dim intRows As Short
        Dim str1 As String
        Dim str2 As String

        Select Case frmName
            Case "frmHome_01"
                str2 = "Report Writer"
            Case "frmStudyDesigner"
                str2 = "Study Designer"
            Case "frmConsole"
                str2 = "StudyDoc Administration"
        End Select

        intRows = Me.cbxModules.Items.Count
        'choose a cbxModule Item
        'choose Report Writer as first item
        For Count1 = 0 To intRows - 1
            str1 = Me.cbxModules.Items(Count1)
            If StrComp(str1, str2, CompareMethod.Text) = 0 Then
                Me.cbxModules.SelectedIndex = Count1
                Exit For
            End If
        Next

        'For Count1 = 0 To intRows - 1
        '    str1 = Me.cbxModulesPers.Items(Count1)
        '    If StrComp(str1, str2, CompareMethod.Text) = 0 Then
        '        Me.cbxModulesPers.SelectedIndex = Count1
        '        Exit For
        '    End If
        'Next

    End Sub

    Sub LoadcbxModules()


        Dim dtbl As System.Data.DataTable
        Dim intRows As Short
        Dim Count1 As Integer
        Dim str1 As String
        Dim cbx As ComboBox = Me.cbxModules
        Dim Count2 As Short

        'dtbl = frmH.tblModules
        'intRows = dtbl.Rows.Count

        cbx.Items.Clear()
        Select Case frmName
            Case "frmConsole"
                cbx.Items.Add("StudyDoc Administration")
                cbx.Items.Add("Report Writer")
            Case "frmHome_01"
                cbx.Items.Add("Report Writer")

        End Select

        cbx.SelectedIndex = 0


    End Sub

    Sub CorporateAddressesInitialize()
        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView
        Dim dtbl1 As System.Data.DataTable
        Dim dtbl2 As System.Data.DataTable
        Dim dtbl3 As System.Data.DataTable
        'Dim dv1 as system.data.dataview
        'Dim dv2 as system.data.dataview
        Dim int2 As Short
        Dim Count1 As Short
        Dim strF As String
        Dim var1
        Dim strS As String
        Dim bool As Boolean


        dtbl1 = tblCorporateAddresses
        'add an unbound column to dtbl3
        bool = dtbl1.Columns.Contains("boolI")
        If bool Then
        Else
            Dim col3 As New DataColumn
            col3.DataType = System.Type.GetType("System.Boolean")
            col3.ColumnName = "boolI"
            col3.Caption = "Active"
            col3.AllowDBNull = True
            dtbl1.Columns.Add(col3)
        End If

        dtbl2 = tblAddressLabels
        dtbl3 = tblCorporateNickNames
        'add an unbound column to dtbl3
        bool = dtbl3.Columns.Contains("boolI")
        If bool Then
        Else
            Dim col3 As New DataColumn
            col3.DataType = System.Type.GetType("System.Boolean")
            col3.ColumnName = "boolI"
            col3.Caption = "Active"
            col3.AllowDBNull = True
            dtbl3.Columns.Add(col3)
        End If

        'update contents of boolI in both tables
        Call UpdateCorpBool()

        dgv1 = Me.dgvNickNames
        dgv1.RowHeadersWidth = 25
        dgv1.AllowUserToOrderColumns = False
        dgv1.AllowUserToResizeColumns = True
        dgv1.AllowUserToResizeRows = True
        dgv1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
        dgv1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

        dgv2 = Me.dgvCorporateAddresses
        dgv2.RowHeadersWidth = 25
        dgv2.AllowUserToOrderColumns = False
        dgv2.AllowUserToResizeColumns = True
        dgv2.AllowUserToResizeRows = True
        dgv2.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
        dgv2.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing


        'fill dgv1 with distinct nicknames
        strF = "charNickName <> '[None]' AND BOOLINCLUDE = -1"
        If Me.rbShowActiveAddresses.Checked Then
            strF = "charNickName <> '[None]' AND BOOLINCLUDE = -1"

        ElseIf Me.rbShowAllAddresses.Checked Then
            strF = "charNickName <> '[None]'"

        ElseIf Me.rbShowInactiveAddresses.Checked Then
            strF = "charNickName <> '[None]' AND BOOLINCLUDE = 0"
        Else
            strF = "charNickName <> '[None]' AND BOOLINCLUDE = -1"
        End If
        Dim dv1 As system.data.dataview = New DataView(dtbl3, strF, "charNickName ASC", DataViewRowState.CurrentRows)
        'dv1.RowFilter = strF
        'dv1.RowStateFilter = DataViewRowState.CurrentRows
        'dv1.Sort = "charNickName ASC"
        dv1.AllowNew = False
        dv1.AllowDelete = False
        int2 = dv1.Count
        dgv1.DataSource = dv1

        'now configure column stuff
        dgv1.AllowUserToResizeColumns = True
        For Count1 = 0 To dgv1.Columns.Count - 1
            'dgv1.Columns(Count1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            dgv1.Columns(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        dgv1.Columns("id_tblCorporateNickNames").Visible = False
        dgv1.Columns("boolI").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv1.Columns("boolI").HeaderText = "Active"
        dgv1.Columns("charNickName").HeaderText = "Nickname"
        dgv1.Columns("charNickName").MinimumWidth = 100
        dgv1.Columns("numincr").Visible = False
        dgv1.Columns("UPSIZE_TS").Visible = False
        dgv1.Columns("boolInclude").Visible = False

        dgv1.RowHeadersWidth = 25

        dgv1.AutoResizeColumns()

        'select first row
        If int2 = 0 Then
        Else
            dgv1.CurrentCell = dgv1.Rows(0).Cells("charNickName")
        End If

        'fill dgv2 with appropriate data
        If int2 = 0 Then
            var1 = 0
        Else
            var1 = dv1(0).Item("id_tblCorporateNickNames")
        End If
        strF = "id_tblCorporateNickNames = " & var1
        strS = "id_tblAddressLabels ASC"

        Dim dv2 As system.data.dataview = New DataView(dtbl1, strF, strS, DataViewRowState.CurrentRows)
        'dv2 = dtbl1.DefaultView
        'dv2.RowFilter = strF
        'dv2.RowStateFilter = DataViewRowState.CurrentRows
        'dv2.Sort = strS
        dv2.AllowNew = False
        dv2.AllowDelete = False
        dgv2.DataSource = dv2

        'configure column stuff
        dgv2.AllowUserToResizeColumns = True

        For Count1 = 0 To dgv2.Columns.Count - 1
            'dgv2.Columns(Count1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            dgv2.Columns(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
            dgv2.Columns(Count1).Visible = False
        Next

        dgv2.Columns("charAddressLabel").Visible = True
        dgv2.Columns("numincr").Visible = False
        dgv2.Columns("charAddressLabel").HeaderText = "Address Label"
        dgv2.Columns("charAddressLabel").ReadOnly = True
        dgv2.Columns("charValue").MinimumWidth = 200
        dgv2.Columns("charValue").Visible = True
        dgv2.Columns("charValue").HeaderText = "Address Value"
        dgv2.Columns("boolI").Visible = True
        dgv2.Columns("boolI").HeaderText = "A *"
        dgv2.Columns("boolI").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv2.Columns("boolI").MinimumWidth = 40
        dgv2.RowHeadersWidth = 25

        dgv2.AutoResizeColumns()


    End Sub

    Sub UpdateCorpBool()
        Dim tbl As System.Data.DataTable
        Dim bool As Boolean
        Dim int1 As Short
        Dim Count1 As Short

        tbl = tblCorporateNickNames
        For Count1 = 0 To tbl.Rows.Count - 1
            int1 = tbl.Rows(Count1).Item("BOOLINCLUDE")
            If int1 = 0 Then
                bool = False
            Else
                bool = True
            End If
            tbl.Rows(Count1).BeginEdit()
            tbl.Rows(Count1).Item("boolI") = bool
            tbl.Rows(Count1).EndEdit()
        Next

        tbl = tblCorporateAddresses
        For Count1 = 0 To tbl.Rows.Count - 1
            int1 = tbl.Rows(Count1).Item("BOOLINCLUDEINTITLE")
            If int1 = 0 Then
                bool = False
            Else
                bool = True
            End If
            tbl.Rows(Count1).BeginEdit()
            tbl.Rows(Count1).Item("boolI") = bool
            tbl.Rows(Count1).EndEdit()
        Next

    End Sub

    Sub CorporateAddressFilter()

        Dim strF As String
        Dim var1
        Dim dv1 As System.Data.DataView
        'Dim dv2 as system.data.dataview
        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView
        Dim dtbl1 As System.Data.DataTable
        Dim dtbl2 As System.Data.DataTable
        Dim intRow As Short
        Dim strS As String
        Dim int1 As Short

        dtbl1 = tblCorporateAddresses
        dtbl2 = tblAddressLabels
        dgv1 = Me.dgvNickNames
        dgv2 = Me.dgvCorporateAddresses

        dv1 = dgv1.DataSource
        int1 = dv1.Count
        If int1 = 0 Then
            Exit Sub
        End If

        intRow = dgv1.CurrentRow.Index
        var1 = dv1(intRow).Item("numincr")

        'fill dgv2 with appropriate data
        If IsDBNull(var1) Then
            var1 = NZ(dv1(intRow).Item("id_tblCorporateNickNames"), "None")
            strF = "id_tblCorporateNickNames = " & var1
        Else
            strF = "numincr = " & var1
        End If
        strS = "id_tblAddressLabels ASC"

        Dim dv2 As System.Data.DataView = New DataView(dtbl1, strF, strS, DataViewRowState.CurrentRows)
        'dv2 = dtbl1.DefaultView
        'dv2.RowFilter = strF
        'dv2.RowStateFilter = DataViewRowState.CurrentRows
        'dv2.Sort = strS
        dv2.AllowNew = False
        dv2.AllowDelete = False
        dgv2.DataSource = dv2

        dgv2.AutoResizeColumns()

    End Sub

    Sub EvalLocks(ByVal rows)


        Dim str1 As String
        Dim ctPB As Short
        Dim ctPBMax As Short
        Dim boolA As Short
        Dim strF As String
        Dim bool As Boolean
        Dim boolE As Boolean

        If Me.chkEditMode.Checked Then
            boolE = True
        Else
            boolE = False
        End If

        'Report Writer

        boolA = BOOLADMINISTRATION ' rows(0).Item("boolAdministration")
        If boolA = 0 Then
            bool = False
        Else
            bool = True
            Call LockAll(Not (bool))
            GoTo end1
        End If

        boolA = BOOLDROPDOWNBOXCONFIGURATION 'rows(0).Item("boolDropdownboxConfiguration")
        If boolA = 0 Then
            bool = False
        Else
            bool = True
        End If
        Call LockDropdownboxTab(Not (bool))

        boolA = BOOLCORPORATEADDRESSES ' rows(0).Item("boolCorporateAddresses")
        If boolA = 0 Then
            bool = False
        Else
            bool = True
        End If
        Call LockCorporateAddressTab(Not (bool))

        boolA = BOOLREPORTTEMPLATEDEFINITIONS 'rows(0).Item("boolReportTemplateDefinitions")
        If boolA = 0 Then
            bool = False
        Else
            bool = True
        End If
        Call LockReportTemplatesTab(Not (bool))

end1:

        'Global Administration

        boolA = BOOLADMINISTRATIONADMIN 'rows(0).Item("boolAdministrationAdmin")
        If boolA = 0 Then
            bool = False
            Call LockAll(Not (bool))
            GoTo end2
        Else
            bool = True

        End If

        boolA = BOOLUSERACCOUNTS 'rows(0).Item("boolUserAccounts")
        If boolA = 0 Then
            bool = False
        Else
            bool = True
        End If
        Call LockUserAccountTab(Not (bool))

        boolA = BOOLHOOKS 'rows(0).Item("boolHooks")
        If boolA = 0 Then
            bool = False
        Else
            bool = True
        End If
        Call LockHooksTab(Not (bool))

        boolA = BOOLCOMPLIANCEGLOBAL ' rows(0).Item("BOOLCOMPLIANCEGLOBAL")
        If boolA = 0 Then
            bool = False
        Else
            bool = True
        End If
        Call LockComplianceTab(Not (bool))

        boolA = BOOLCUSTOMFIELDCODES 'rows(0).Item("BOOLCUSTOMFIELDCODES")
        If boolA = 0 Then
            bool = False
        Else
            bool = True
        End If
        Call LockFCTab(Not (bool))

        boolA = BOOLPERMISSIONSMANAGER 'rows(0).Item("BOOLPERMISSIONSMANAGER")
        If boolA = 0 Then
            bool = False
        Else
            bool = True
        End If
        Call LockPM(Not (bool))

        str1 = Me.cbxModules.SelectedItem
        Select Case str1
            Case "StudyDoc Administration"
                boolA = BOOLGLOBALPARAMETERSADMIN 'rows(0).Item("BOOLGLOBALPARAMETERSADMIN")
                If boolA = 0 Then
                    bool = False
                Else
                    bool = True
                End If
                Call LockGlobalParametersAdmin(Not (bool))
            Case "Report Writer"
                boolA = BOOLGLOBALPARAMETERS ' rows(0).Item("boolGlobalParameters")
                If boolA = 0 Then
                    bool = False
                Else
                    bool = True
                End If
                Call LockGlobalTab(Not (bool))
            Case "Study Designer"
        End Select

        'show lblrestrictions
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String

        str2 = Me.lbxTab1.SelectedItem
        str3 = Me.cbxModules.SelectedItem
        Select Case str3
            Case "StudyDoc Administration"
                boolA = BOOLADMINISTRATIONADMIN 'rows(0).Item("boolAdministrationAdmin")
                If boolA = 0 Then
                    bool = False
                    Call LockAll(Not (bool))
                    GoTo end3
                Else
                    bool = True

                End If

                Select Case str2
                    Case "User Accounts"
                        boolA = BOOLUSERACCOUNTS ' rows(0).Item("boolUserAccounts")
                        If boolA = 0 Then
                            bool = False
                        Else
                            bool = True
                        End If
                    Case "Global Parameters"
                        str4 = "boolGlobalParameters"
                        Select Case str3
                            Case "StudyDoc Administration"
                                boolA = BOOLGLOBALPARAMETERSADMIN ' rows(0).Item(str4)
                                If boolA = 0 Then
                                    bool = False
                                Else
                                    bool = True
                                End If
                                Call LockGlobalParametersAdmin(Not (bool))
                            Case "Report Writer"
                                boolA = BOOLGLOBALPARAMETERS ' rows(0).Item(str4)
                                If boolA = 0 Then
                                    bool = False
                                Else
                                    bool = True
                                End If
                                Call LockGlobalTab(Not (bool))
                            Case "Study Designer"
                                'tbd

                        End Select

                    Case "Hooks"
                        boolA = BOOLHOOKS 'rows(0).Item("boolHooks")
                        If boolA = 0 Then
                            bool = False
                        Else
                            bool = True
                        End If
                        Call LockHooksTab(Not (bool))
                    Case "Permissions Manager"
                        boolA = BOOLPERMISSIONSMANAGER ' rows(0).Item("BOOLPERMISSIONSMANAGER")
                        If boolA = 0 Then
                            bool = False
                        Else
                            bool = True
                        End If
                        Call LockPermmissionsTab(Not (bool))

                End Select
end3:
            Case "Report Writer"

                boolA = BOOLADMINISTRATION 'rows(0).Item("boolAdministration")
                If boolA = 0 Then
                    bool = False
                    Call LockAll(Not (bool))
                    GoTo end4
                Else
                    bool = True

                End If

                Select Case str2
                    Case "Dropdownbox Configuration"
                        boolA = BOOLDROPDOWNBOXCONFIGURATION 'rows(0).Item("boolDropdownboxConfiguration")
                        If boolA = 0 Then
                            bool = False
                        Else
                            bool = True
                        End If
                    Case "Corporate Addresses"
                        boolA = BOOLCORPORATEADDRESSES 'rows(0).Item("boolCorporateAddresses")
                        If boolA = 0 Then
                            bool = False
                        Else
                            bool = True
                        End If
                    Case "Study Template Definitions"
                        boolA = BOOLREPORTTEMPLATEDEFINITIONS ' rows(0).Item("boolReportTemplateDefinitions")
                        If boolA = 0 Then
                            bool = False
                        Else
                            bool = True
                        End If
                    Case "Global Parameters"
                        str4 = "boolGlobalParameters"
                        Select Case str3
                            Case "StudyDoc Administration"

                                boolA = BOOLGLOBALPARAMETERSADMIN ' rows(0).Item(str4)
                                If boolA = 0 Then
                                    bool = False
                                Else
                                    bool = True
                                End If
                                Call LockGlobalParametersAdmin(Not (bool))
                            Case "Report Writer"

                                boolA = BOOLGLOBALPARAMETERS ' rows(0).Item(str4)
                                If boolA = 0 Then
                                    bool = False
                                Else
                                    bool = True
                                End If
                                Call LockGlobalTab(Not (bool))
                            Case "Study Designer"
                        End Select

                    Case "Hooks"
                        boolA = BOOLHOOKS 'rows(0).Item("boolHooks")
                        If boolA = 0 Then
                            bool = False
                        Else
                            bool = True
                        End If
                        Call LockHooksTab(Not (bool))
                    Case "Custom Field Codes"
                        boolA = BOOLCUSTOMFIELDCODES 'rows(0).item("boolCustomFieldCodes")
                        If boolA = 0 Then
                            bool = False
                        Else
                            bool = True
                        End If
                        Me.panFCcmd.Enabled = bool
                        Call LockCustomFieldCodeTab(Not (bool))

                End Select
end4:
            Case "Study Designer"
                bool = False
        End Select

end2:
        'Me.lblRestricted.Visible = Not (bool)

        If boolE Then
        Else
            Call LockAll(True)
        End If

    End Sub

    Sub UserIDActions()

        If Me.cmdSave.Enabled Then 'continue

            Dim dgvU As DataGridView = Me.dgvUsers
            Dim intRowU As Int16 = Me.dgvUsers.CurrentRow.Index

            Dim dgv As DataGridView
            Dim var1

            dgv = Me.dgvUserAttributes


            Dim idP As Int64
            Dim intRow As Int32
            Try
                intRow = Me.dgvUsers.CurrentRow.Index
                idP = Me.dgvUsers("ID_TBLPERSONNEL", intRow).Value
            Catch ex As Exception
                idP = 0
            End Try

            If dgv.RowCount = 0 Then
                Me.gbxPassword.Enabled = False
                'Me.cmdAddUser.Enabled = False
                Me.cmdEnterPassword.Enabled = False
                Me.gbWindowsAuth.Enabled = False
                Me.gbWatsonAccount.Enabled = False
                Me.gbSetPerm.Enabled = False
            Else

                Me.cmdEnterPassword.Enabled = True
                If intRowU = 0 Then
                    Me.gbWindowsAuth.Enabled = False
                    Me.gbWatsonAccount.Enabled = False
                Else
                    Me.gbWindowsAuth.Enabled = True
                    Me.gbWatsonAccount.Enabled = True
                End If


                If idP = 1 Then
                    Me.gbxPassword.Enabled = False
                    Me.cmdAddUserID.Enabled = False
                    Me.gbSetPerm.Enabled = False
                Else
                    Me.gbxPassword.Enabled = True
                    Me.cmdAddUserID.Enabled = True
                    Me.gbSetPerm.Enabled = True
                End If
            End If

        Else
            Me.gbxPassword.Enabled = False
            Me.cmdAddUser.Enabled = False
            Me.cmdEnterPassword.Enabled = False
            Me.gbWindowsAuth.Enabled = False
            Me.gbWatsonAccount.Enabled = False
            Me.gbSetPerm.Enabled = False
        End If


    End Sub

    Sub DoThisAdmin(ByVal cmd As String)

        'Dim frm As New frmProgress_01
        Dim str1 As String
        Dim ctPB As Short
        Dim ctPBMax As Short
        Dim boolA As Short
        Dim var1, var2
        Dim Count1 As Int32
        Dim boolFL As Boolean

        boolAddRow = True

        Cursor.Current = Cursors.WaitCursor

        Dim strF As String
        Dim bool As Boolean

        strF = "ID_TBLPERMISSIONS = " & id_tblPermissions
        Dim rows() As DataRow
        rows = tblPermissions.Select(strF)
        If rows.Length = 0 Then
            MsgBox("Guest does not have Edit privileges.", MsgBoxStyle.Information, "No no...")
            Exit Sub
        Else
        End If

        Select Case cmd
            Case "Edit"

                ReDim arrPswd(10, 100)
                ctPswd = 0

                Me.chkEditMode.Checked = True

                Me.cbxModules.Enabled = False

                Call EvalLocks(rows)

                SetToEditMode()

            Case "Save"

                Dim tUserID As String
                Dim tUserName As String

                tUserID = gUserID
                tUserName = gUserName

                strRFC = GetDefaultRFC()
                strMOS = GetDefaultMOS()

                gATAdds = 0
                gATDeletes = 0
                gATMods = 0

                Me.dgvPermissions.Enabled = True

                Dim boolESTemp As Boolean = Me.rbESigOn.Checked
                Dim boolT As Boolean = Me.rbAuditTrailOn.Checked

                If boolT Then
                    gboolAuditTrail = True
                    If boolESTemp Then
                        gboolESig = True
                    End If
                End If

                'If (gboolAuditTrail And gboolESig) Or boolESigTemp <> gboolESig Or boolESTemp <> gboolESig Then
                If gboolAuditTrail And gboolESig Then ' Or boolESigTemp <> gboolESig Or boolESTemp <> gboolESig Then

                    Dim frm As New frmESig

                    frm.ShowDialog()

                    If frm.boolCancel Then
                        frm.Dispose()
                        Exit Sub
                    End If

                    'NO! User must be current logon
                    'gUserID = frm.tUserID
                    'gUserName = frm.tUserName

                    frm.Dispose()

                End If

                'clear audittrailtemp
                tblAuditTrailTemp.Clear()
                idSE = 0

                Dim dt1 As DateTime
                dt1 = Now


                Me.chkEditMode.Checked = False

                Me.cbxModules.Enabled = True

                str1 = "Saving data..."
                Call PositionProgress()
                Me.lblProgress.Text = str1
                Me.lblProgress.Visible = True
                Me.lblProgress.Refresh()
                ctPB = 0
                ctPBMax = 6
                Me.pb1.Value = 1
                Me.pb1.Maximum = ctPBMax
                Me.pb1.Visible = True
                Me.pb1.Refresh()
                Me.Refresh()

                boolA = BOOLUSERACCOUNTS
                If boolA = 0 Then
                    bool = False
                Else
                    bool = True
                End If
                If bool Then
                    ctPB = ctPB + 1
                    Me.pb1.Value = ctPB
                    Me.pb1.Refresh()
                    Call SaveUserAccountTab(dt1)
                    Call SaveLDAP()
                Else
                End If

                boolA = BOOLPERMISSIONSMANAGER
                If boolA = 0 Then
                    bool = False
                Else
                    bool = True
                End If
                If bool Then
                    ctPB = ctPB + 1
                    Me.pb1.Value = ctPB
                    Me.pb1.Refresh()
                    Call SavePermissionsGroup(dt1)
                Else
                End If

                boolA = BOOLDROPDOWNBOXCONFIGURATION
                If boolA = 0 Then
                    bool = False
                Else
                    bool = True
                End If
                If bool Then
                    ctPB = ctPB + 1
                    Me.pb1.Value = ctPB
                    Me.pb1.Refresh()
                    Call SaveDropdownboxTab()
                Else
                End If

                boolA = BOOLCORPORATEADDRESSES
                If boolA = 0 Then
                    bool = False
                Else
                    bool = True
                End If
                If bool Then
                    ctPB = ctPB + 1
                    Me.pb1.Value = ctPB
                    Me.pb1.Refresh()
                    Call SaveCorporateAddressTab()
                Else
                End If

                boolA = BOOLREPORTTEMPLATEDEFINITIONS
                If boolA = 0 Then
                    bool = False
                Else
                    bool = True
                End If
                If bool Then
                    ctPB = ctPB + 1
                    Me.pb1.Value = ctPB
                    Me.pb1.Refresh()
                    Call SaveReportTemplatesTab()
                Else
                End If

                boolA = BOOLGLOBALPARAMETERS
                If boolA = 0 Then
                    bool = False
                Else
                    bool = True
                End If
                If bool Then
                    ctPB = ctPB + 1
                    Me.pb1.Value = ctPB
                    Me.pb1.Refresh()
                    Call SaveGlobalTab()
                    'Call SaveAdminFC()?????
                Else
                End If

                boolA = BOOLHOOKS
                If boolA = 0 Then
                    bool = False
                Else
                    bool = True
                End If
                If bool Then
                    Call SaveHooksTab()
                End If

                boolA = BOOLCOMPLIANCEGLOBAL
                If boolA = 0 Then
                    bool = False
                Else
                    bool = True
                End If
                If bool Then
                    Call SaveComplianceTab()
                End If

                boolA = BOOLCUSTOMFIELDCODES
                If boolA = 0 Then
                    bool = False
                Else
                    bool = True
                End If
                If bool Then
                    Call SaveAdminFC()
                End If

                'record tblaudittrailtemp
                'dt1 = RecordAuditTrail(True, dt1)
                'what am i doing here?
                'If Me.rbAuditTrailOn.Checked Then
                '    gboolAuditTrail = True
                'Else
                '    gboolAuditTrail = False
                'End If
                'If gboolAuditTrail <> boolAuditTrailTemp Then
                '    boolT = gboolAuditTrail
                '    gboolAuditTrail = True
                'End If
                Call RecordAuditTrail(True, dt1)

                If boolT Then
                    gboolAuditTrail = True
                    If boolESTemp Then
                        gboolESig = True
                    Else
                        gboolESig = False
                    End If
                Else
                    boolAuditTrail = False
                    gboolESig = False
                End If

                ''NO! 
                'gUserID = tUserID
                'gUserName = tUserName

                Dim idUID As Int64
                Dim strPswd As String
                For Count1 = 1 To ctPswd
                    idUID = arrPswd(2, Count1)
                    strPswd = arrPswd(3, Count1)
                    Call SavePswdHistory(idUID, strPswd, dt1)
                Next

                Call LockUserAccountTab(True)
                Call LockDropdownboxTab(True)
                Call LockCorporateAddressTab(True)
                Call LockReportTemplatesTab(True)
                Call LockGlobalTab(True)
                Call LockHooksTab(True)
                Call LockComplianceTab(True)
                Call LockFCTab(True)
                Call LockPM(True)

                Me.lblProgress.Visible = False
                Me.pb1.Visible = False
                Me.Refresh()

                SetToNonEditMode()

                're-populate dropdownboxes
                Call FillDataCbx()
                'boolHomeCBox = True 'after fill dropboxes, must re-assign the datagridviewcomboboxcells
                Call FillDropdownBoxes()
                Call FillCorporateNames()

                Call FillUserboolA(Me.dgvUsers)
                Call FillUserboolA(Me.dgvUserAttributes)
                Call FillFCRW()

                Call FillcbxPermissionsGroup(True)


                Call LockAll(True)

                Call SetPermissions(True)

            Case "Cancel"

                Me.dgvPermissions.Enabled = True

                Me.chkEditMode.Checked = False

                Me.cbxModules.Enabled = True

                Call DoCancelUserAccountTab()
                Call DoCancelDropdownboxTab()
                Call DoCancelCorporateAddressTab()
                Call DoCancelReportTemplatesTab()

                boolFL = boolFormLoad
                boolFormLoad = True
                Call DoCancelPermissionsManager()
                boolFormLoad = boolFL

                Call DoCancelGlobalTab()
                Call DoCancelHooksTab()
                Call DoCancelComplianceTab()
                Call DoFCAdminCancel()
                Call LockUserAccountTab(True)
                Call LockDropdownboxTab(True)
                Call LockCorporateAddressTab(True)
                Call LockReportTemplatesTab(True)
                Call LockGlobalTab(True)
                Call LockHooksTab(True)
                Call LockComplianceTab(True)
                Call LockFCTab(True)
                Call LockPM(True)

                SetToNonEditMode()

                Call FillUserboolA(Me.dgvUsers)
                Call FillUserboolA(Me.dgvUserAttributes)

                Call FilllvPermissions()

        End Select

        Call EvaluateUserAccounts() 'call this to get appropriate setting for cmdEnterPassword

        Call UserIDActions()

        Call SetLDAP()

        Call LDAPDisplaySettings()



        boolAddRow = False
        Cursor.Current = Cursors.Default

    End Sub



    Sub SavePermissionsGroup(dt1 As DateTime)

        Me.dgvPermissions.CommitEdit(DataGridViewDataErrorContexts.Commit)

        Call FillAuditTrailTemp(tblPermissions)

        If boolGuWuOracle Then
            Try
                ta_tblPermissions.Update(tblPermissions)
            Catch ex As DBConcurrencyException
                'ds2005.TBLPERMISSIONS.Merge('ds2005.TBLPERMISSIONS, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_tblPermissionsAcc.Update(tblPermissions)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLPERMISSIONS.Merge('ds2005Acc.TBLPERMISSIONS, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_tblPermissionsSQLServer.Update(tblPermissions)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLPERMISSIONS.Merge('ds2005Acc.TBLPERMISSIONS, True)
            End Try
        End If

        Me.dgvPermissions.Enabled = True

    End Sub

    Sub SaveUserAccountTab(dt1 As DateTime)

        Dim int1 As Short
        Dim int2 As Short
        Dim intU As Short
        Dim intA As Short

        'record current row
        If Me.dgvUserAttributes.Rows.Count = 0 Then
            intA = -1
        Else
            intA = Me.dgvUserAttributes.CurrentRow.Index
        End If

        If Me.dgvUsers.CurrentRow Is Nothing Then
            int1 = -1
        Else
            int1 = Me.dgvUsers.CurrentRow.Index
            'record current row
        End If
        intU = int1



        Me.dgvUsers.CommitEdit(DataGridViewDataErrorContexts.Commit)
        Me.dgvUserAttributes.CommitEdit(DataGridViewDataErrorContexts.Commit)

        Call FillAuditTrailTemp(tblPersonnel)
        Call FillAuditTrailTemp(tblUserAccounts)


        Dim arrUA()
        Dim arrUID()

        Dim intUA As Int32 = 0
        Dim intUID As Int32 = 0

        Dim tblUA As System.Data.DataTable
        Dim tblUID As System.Data.DataTable
        Dim intRowsUA As Int16 = 0
        Dim intRowsUID As Int16 = 0
        Dim dr As DataRow
        Dim rowsD() As DataRow
        Dim Count1 As Int32
        Dim var1
        Dim strF As String
        Dim strF1 As String


        'these items are not updating
        tblUA = tblPersonnel.GetChanges(DataRowState.Added)
        tblUID = tblUserAccounts.GetChanges(DataRowState.Added)
        'I'm getting PO'd
        'try saving ID's and add later

        'update tblUserAcct and tblPersonnel
        If tblUA Is Nothing Then
        Else
            ReDim arrUA(tblUA.Rows.Count)
            For Count1 = 0 To tblUA.Rows.Count - 1
                intUA = intUA + 1
                arrUA(intUA) = tblUA.Rows(Count1).Item("ID_TBLPERSONNEL")
            Next
        End If

        If tblUID Is Nothing Then
        Else
            ReDim arrUID(tblUID.Rows.Count)
            For Count1 = 0 To tblUID.Rows.Count - 1
                intUID = intUID + 1
                arrUID(intUID) = tblUID.Rows(Count1).Item("ID_TBLUSERACCOUNTS")
            Next
        End If


        '*****

        'for some reason, user attributes aren't updating
        'TRY UPDATING

        If intUA = 0 Then
        Else
            For Count1 = 1 To intUA
                var1 = arrUA(Count1)
                Dim rowsI() As DataRow
                strF = "ID_TBLPERSONNEL = " & var1
                rowsI = tblPersonnel.Select(strF)
                rowsI(0).BeginEdit()
                rowsI(0).Item("UPSIZE_TS") = dt1
                rowsI(0).Item("DTACTIVATED") = dt1
                rowsI(0).EndEdit()
            Next
            'tblPermissions.AcceptChanges()
        End If

        If intUID = 0 Then
        Else
            For Count1 = 1 To intUID
                var1 = arrUID(Count1)
                Dim rowsI() As DataRow
                strF = "ID_TBLUSERACCOUNTS = " & var1
                rowsI = tblUserAccounts.Select(strF)
                rowsI(0).BeginEdit()
                rowsI(0).Item("DTTIMESTAMP") = dt1
                rowsI(0).Item("UPSIZE_TS") = dt1
                rowsI(0).Item("DTACTIVATED") = dt1
                rowsI(0).EndEdit()
            Next
            'tblUserAccounts.AcceptChanges()
        End If

        'do  tblWatsonUsers


        '****

        If boolGuWuOracle Then
            Try
                ta_tblPersonnel.Update(tblPersonnel)
            Catch ex As DBConcurrencyException
                'ds2005.TBLPERSONNEL.Merge('ds2005.TBLPERSONNEL, True)
            End Try

            Try
                ta_tblUserAccounts.Update(tblUserAccounts)
            Catch ex As DBConcurrencyException
                'ds2005.TBLUSERACCOUNTS.Merge('ds2005.TBLUSERACCOUNTS, True)
            End Try

        ElseIf boolGuWuAccess Then
            Try
                ta_tblPersonnelAcc.Update(tblPersonnel)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLPERSONNEL.Merge('ds2005Acc.TBLPERSONNEL, True)
            End Try

            Try
                ta_tblUserAccountsAcc.Update(tblUserAccounts)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLUSERACCOUNTS.Merge('ds2005Acc.TBLUSERACCOUNTS, True)
            End Try

        ElseIf boolGuWuSQLServer Then
            Try
                ta_tblPersonnelSQLServer.Update(tblPersonnel)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLPERSONNEL.Merge('ds2005Acc.TBLPERSONNEL, True)
            End Try

            Try
                ta_tblUserAccountsSQLServer.Update(tblUserAccounts)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLUSERACCOUNTS.Merge('ds2005Acc.TBLUSERACCOUNTS, True)
            End Try

        End If

        'reset stuff
        Call ConfigureUserAccountAttributes(True)

        int2 = Me.dgvUsers.Rows.Count
        Try
            If intU > int2 - 1 Then
                Call ShowUsers(0)
            Else
                Call ShowUsers(intU)
            End If
        Catch ex As Exception

        End Try


        int2 = Me.dgvUserAttributes.Rows.Count
        Try
            If intA > int2 - 1 Then
                Call ShowAccounts(0)
            Else
                Call ShowAccounts(intA)
            End If
        Catch ex As Exception

        End Try


        'stuff isn't setting correctly for some reason
        'Call ConfigureUserAccountAttributes()
        Try
            Call FillUserboolA(Me.dgvUsers)
        Catch ex As Exception

        End Try

        Try
            Call FillUserboolA(Me.dgvUserAttributes)
        Catch ex As Exception

        End Try


    End Sub

    Sub SaveDropdownboxTab()

        Dim intRow As Short
        If Me.dgvDropdownboxContents.CurrentRow Is Nothing Then
            intRow = -1
        Else
            intRow = Me.dgvDropdownboxContents.CurrentRow.Index
        End If

        Me.dgvDropdownboxContents.CommitEdit(DataGridViewDataErrorContexts.Commit)

        Call FillAuditTrailTemp(tblDropdownBoxContent)

        If boolGuWuOracle Then
            Try
                ta_tblDropdownBoxContent.Update(tblDropdownBoxContent)
            Catch ex As DBConcurrencyException
                'ds2005.TBLDROPDOWNBOXCONTENT.Merge('ds2005.TBLDROPDOWNBOXCONTENT, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_tblDropdownBoxContentAcc.Update(tblDropdownBoxContent)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLDROPDOWNBOXCONTENT.Merge('ds2005Acc.TBLDROPDOWNBOXCONTENT, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_tblDropdownBoxContentSQLServer.Update(tblDropdownBoxContent)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLDROPDOWNBOXCONTENT.Merge('ds2005Acc.TBLDROPDOWNBOXCONTENT, True)
            End Try
        End If

        If intRow = -1 Then
        Else
            Me.dgvDropdownboxContents.Rows(intRow).Cells("charValue").Selected = True
        End If

        Call DropDownsConfigure()

    End Sub

    Sub SaveCompliance(ByVal boolTemp As Boolean)

        If boolFormLoad Then
            Exit Sub
        End If

        Dim dtbl As System.Data.DataTable = tblConfigCompliance
        Dim var1

        dtbl.Rows(0).BeginEdit()

        If Me.rbAuditTrailOn.Checked Then
            var1 = -1
            'If boolTemp Then
            'Else
            '    gboolAuditTrail = True
            'End If
        Else
            var1 = 0
            'If boolTemp Then
            'Else
            '    gboolAuditTrail = False
            'End If
        End If
        dtbl.Rows(0).Item("BOOLAUDITTRAIL") = var1

        If Me.rbESigOn.Checked Then
            var1 = -1
            'If boolTemp Then
            'Else
            '    gboolESig = True
            'End If
        Else
            var1 = 0
            'If boolTemp Then
            'Else
            '    gboolESig = False
            'End If
        End If
        dtbl.Rows(0).Item("BOOLESIG") = var1

        If Me.rbOnlyLoggedOn.Checked Then
            var1 = -1
        Else
            var1 = 0
        End If
        dtbl.Rows(0).Item("BOOLLOGGEDONUSER") = var1

        If Me.chkMeaningOfSign.Checked Then
            var1 = -1
        Else
            var1 = 0
        End If
        dtbl.Rows(0).Item("BOOLMEANINGOFSIG") = var1

        If Me.chkSigFreeForm.Checked Then
            var1 = -1
        Else
            var1 = 0
        End If
        dtbl.Rows(0).Item("BOOLRESTRICTSIG") = var1

        If Me.chkReasonForChange.Checked Then
            var1 = -1
        Else
            var1 = 0
        End If
        dtbl.Rows(0).Item("BOOLREASONFORCHANGE") = var1

        If Me.chkReasonFreeForm.Checked Then
            var1 = -1
        Else
            var1 = 0
        End If
        dtbl.Rows(0).Item("BOOLRESTRICTREASON") = var1

        dtbl.Rows(0).EndEdit()

    End Sub

    Sub SaveComplianceTab()

        'Call SaveCompliance(False)
        Call SaveCompliance(True)

        Dim dgv1 As DataGridView = Me.dgvMOS
        Dim dgv2 As DataGridView = Me.dgvRFC

        dgv1.CommitEdit(DataGridViewDataErrorContexts.Commit)
        dgv2.CommitEdit(DataGridViewDataErrorContexts.Commit)

        'Dim boolT As Boolean = gboolAuditTrail
        'If gboolAuditTrail <> boolAuditTrailTemp Then
        '    boolT = gboolAuditTrail
        '    gboolAuditTrail = True
        'End If
        Call FillAuditTrailTemp(tblConfigCompliance)
        'gboolAuditTrail = boolT

        If boolGuWuOracle Then
            Try
                ta_tblConfigCompliance.Update(tblConfigCompliance)
            Catch ex As DBConcurrencyException
                'ds2005.tblConfigCompliance.Merge('ds2005.tblConfigCompliance, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_tblConfigComplianceAcc.Update(tblConfigCompliance)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLCONFIGCOMPLIANCE.Merge('ds2005Acc.TBLCONFIGCOMPLIANCE, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_tblConfigComplianceSQLServer.Update(tblConfigCompliance)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLCONFIGCOMPLIANCE.Merge('ds2005Acc.TBLCONFIGCOMPLIANCE, True)
            End Try
        End If

        Call FillAuditTrailTemp(tblMeaningOfSig)

        If boolGuWuOracle Then
            Try
                ta_tblMeaningOfSig.Update(tblMeaningOfSig)
            Catch ex As DBConcurrencyException
                'ds2005.tblMeaningOfSig.Merge('ds2005.tblMeaningOfSig, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_tblMeaningOfSigAcc.Update(tblMeaningOfSig)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLMEANINGOFSIG.Merge('ds2005Acc.TBLMEANINGOFSIG, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_tblMeaningOfSigSQLServer.Update(tblMeaningOfSig)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLMEANINGOFSIG.Merge('ds2005Acc.TBLMEANINGOFSIG, True)
            End Try
        End If

        Call FillAuditTrailTemp(tblReasonForChange)

        If boolGuWuOracle Then
            Try
                ta_tblReasonForChange.Update(tblReasonForChange)
            Catch ex As DBConcurrencyException
                'ds2005.tblReasonForChange.Merge('ds2005.tblReasonForChange, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_tblReasonForChangeAcc.Update(tblReasonForChange)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLREASONFORCHANGE.Merge('ds2005Acc.TBLREASONFORCHANGE, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_tblReasonForChangeSQLServer.Update(tblReasonForChange)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLREASONFORCHANGE.Merge('ds2005Acc.TBLREASONFORCHANGE, True)
            End Try
        End If

    End Sub

    Sub SaveHooksTab()

        Me.dgvHooks.CommitEdit(DataGridViewDataErrorContexts.Commit)

        Call FillAuditTrailTemp(tblHooks)

        If boolGuWuOracle Then
            Try
                ta_tblHooks.Update(tblHooks)
            Catch ex As DBConcurrencyException
                'ds2005.TBLHOOKS.Merge('ds2005.TBLHOOKS, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_tblHooksAcc.Update(tblHooks)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLHOOKS.Merge('ds2005Acc.TBLHOOKS, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_tblHooksSQLServer.Update(tblHooks)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLHOOKS.Merge('ds2005Acc.TBLHOOKS, True)
            End Try
        End If

        'fill unbound column values
        Call UpdateHookActive()

        Me.dgvHooks.AutoResizeColumns()


    End Sub

    Sub SaveGlobalTab()

        Dim intRow As Short
        If Me.dgvGlobal.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = Me.dgvGlobal.CurrentRow.Index
        End If

        Me.dgvGlobal.CommitEdit(DataGridViewDataErrorContexts.Commit)

        Call FillAuditTrailTemp(tblConfiguration)

        If boolGuWuOracle Then
            Try
                ta_tblConfiguration.Update(tblConfiguration)
            Catch ex As DBConcurrencyException
                'ds2005.TBLCONFIGURATION.Merge('ds2005.TBLCONFIGURATION, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_tblConfigurationAcc.Update(tblConfiguration)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLCONFIGURATION.Merge('ds2005Acc.TBLCONFIGURATION, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_tblConfigurationSQLServer.Update(tblConfiguration)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLCONFIGURATION.Merge('ds2005Acc.TBLCONFIGURATION, True)
            End Try
        End If


        'update global variables
        'Gdateformat
        Dim rows1() As DataRow
        Dim str1 As String
        Dim var1
        Dim tbl1 As System.Data.DataTable

        tbl1 = tblConfiguration
        str1 = "CHARCONFIGGROUP = 'Global' AND CHARCONFIGCATEGORY = 'Global Settings' AND CHARCONFIGTITLE = 'Table Date Format'"
        rows1 = tbl1.Select(str1)
        var1 = NZ(rows1(0).Item("CHARCONFIGVALUE"), "")
        If Len(var1) = 0 Then
            GDateFormat = "MM/dd/yyyy"
        Else
            GDateFormat = var1
        End If

        str1 = "CHARCONFIGGROUP = 'Global' AND CHARCONFIGCATEGORY = 'Global Settings' AND CHARCONFIGTITLE = 'Text Date Format'"
        Erase rows1
        rows1 = tbl1.Select(str1)
        var1 = NZ(rows1(0).Item("CHARCONFIGVALUE"), "")
        If Len(var1) = 0 Then
            GTextDateFormat = "MMMM dd, yyyy"
        Else
            GTextDateFormat = var1
        End If


        'GSigFig
        Erase rows1
        str1 = "CHARCONFIGGROUP = 'Global' AND CHARCONFIGCATEGORY = 'Global Settings' AND CHARCONFIGTITLE = 'Default Significant Figures for Conc Data'"
        rows1 = tbl1.Select(str1)
        var1 = NZ(rows1(0).Item("CHARCONFIGVALUE"), "")
        If Len(var1) = 0 Then
            GSigFig = 3
        Else
            GSigFig = var1
        End If

        'GDec
        Erase rows1
        str1 = "CHARCONFIGGROUP = 'Global' AND CHARCONFIGCATEGORY = 'Global Settings' AND CHARCONFIGTITLE = 'Default # of Decimals for Conc Data'"
        rows1 = tbl1.Select(str1)
        var1 = NZ(rows1(0).Item("CHARCONFIGVALUE"), "")
        If Len(var1) = 0 Then
            GDec = 1
        Else
            GDec = var1
        End If

        'GTimeZone
        Erase rows1
        str1 = "CHARCONFIGGROUP = 'Global' AND CHARCONFIGCATEGORY = 'Global Settings' AND CHARCONFIGTITLE = 'Default Time Zone'"
        rows1 = tbl1.Select(str1)
        var1 = NZ(rows1(0).Item("CHARCONFIGVALUE"), "")
        If Len(var1) = 0 Then
            GTimeZone = "Eastern Time Zone"
        Else
            GTimeZone = var1
        End If

        'GIncDiff
        Erase rows1
        str1 = "CHARCONFIGGROUP = 'Global' AND CHARCONFIGCATEGORY = 'Global Settings' AND CHARCONFIGTITLE = 'Default Incurred Sample %Diff Calculation'"
        rows1 = tbl1.Select(str1)
        var1 = NZ(rows1(0).Item("CHARCONFIGVALUE"), "")
        If Len(var1) = 0 Then
            GIncDiff = "%Difference"
        Else
            GIncDiff = var1
        End If

        ''boolUseSigFigs
        Erase rows1
        str1 = "CHARCONFIGGROUP = 'Global' AND CHARCONFIGCATEGORY = 'Global Settings' AND CHARCONFIGTITLE = 'Default Use Significant Figures, not Decimals for Conc Data'"
        rows1 = tbl1.Select(str1)
        var1 = NZ(rows1(0).Item("CHARCONFIGVALUE"), "")
        If Len(var1) = 0 Then
            boolGUseSigFigs = True
        Else
            If StrComp(var1, "FALSE", CompareMethod.Text) = 0 Then
                boolGUseSigFigs = False
            Else
                boolGUseSigFigs = True
            End If
        End If

        'boolUseHyperlinks
        Erase rows1
        str1 = "CHARCONFIGGROUP = 'Global' AND CHARCONFIGCATEGORY = 'Global Settings' AND CHARCONFIGTITLE = 'Use Hyperlink Feature'"
        rows1 = tbl1.Select(str1)
        var1 = NZ(rows1(0).Item("CHARCONFIGVALUE"), "")
        If Len(var1) = 0 Then
            boolUseHyperlinks = True
        Else
            If StrComp(var1, "False", CompareMethod.Text) = 0 Then
                boolUseHyperlinks = False
            Else
                boolUseHyperlinks = True
            End If
        End If

        'gintQCDec
        str1 = "CHARCONFIGGROUP = 'Global' AND CHARCONFIGCATEGORY = 'Global Settings' AND CHARCONFIGTITLE = 'Default # of Decimals for QC Stats'"
        rows1 = tbl1.Select(str1)
        var1 = NZ(rows1(0).Item("CHARCONFIGVALUE"), 1)
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

        'gGoToWord
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

        'update panel in frmH after gboolER update
        Call ConfigLockFinalReport()

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

        Me.dgvGlobal.Rows(intRow).Cells("charConfigValue").Selected = True

    End Sub

    Sub SaveCorporateAddressTab()

        Dim intRow As Short
        If Me.dgvNickNames.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = Me.dgvNickNames.CurrentRow.Index
        End If

        Me.dgvCorporateAddresses.CommitEdit(DataGridViewDataErrorContexts.Commit)
        Me.dgvNickNames.CommitEdit(DataGridViewDataErrorContexts.Commit)

        Call FillAuditTrailTemp(tblCorporateNickNames)
        Call FillAuditTrailTemp(tblCorporateAddresses)

        If boolGuWuOracle Then
            Try
                ta_tblCorporateNickNames.Update(tblCorporateNickNames)
            Catch ex As DBConcurrencyException
                'ds2005.TBLCORPORATENICKNAMES.Merge('ds2005.TBLCORPORATENICKNAMES, True)
            End Try

            Try
                ta_tblCorporateAddresses.Update(tblCorporateAddresses)
            Catch ex As DBConcurrencyException
                'ds2005.TBLCORPORATEADDRESSES.Merge('ds2005.TBLCORPORATEADDRESSES, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_tblCorporateNickNamesAcc.Update(tblCorporateNickNames)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLCORPORATENICKNAMES.Merge('ds2005Acc.TBLCORPORATENICKNAMES, True)
            End Try

            Try
                ta_tblCorporateAddressesAcc.Update(tblCorporateAddresses)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLCORPORATEADDRESSES.Merge('ds2005Acc.TBLCORPORATEADDRESSES, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_tblCorporateNickNamesSQLServer.Update(tblCorporateNickNames)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLCORPORATENICKNAMES.Merge('ds2005Acc.TBLCORPORATENICKNAMES, True)
            End Try

            Try
                ta_tblCorporateAddressesSQLServer.Update(tblCorporateAddresses)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLCORPORATEADDRESSES.Merge('ds2005Acc.TBLCORPORATEADDRESSES, True)
            End Try
        End If

        Call CorporateAddressesInitialize()
        Call CorporateAddressFilter()

        'select row
        If Me.dgvNickNames.Rows.Count = 0 Then
        Else
            Me.dgvNickNames.CurrentCell = Me.dgvNickNames.Rows(intRow).Cells("charNickname")
        End If

        Call ShowAddresses()


        'populate dropdown boxes on frmHome
        Call FillCorporateNames()


    End Sub

    Sub SaveReportTemplatesTab()
        Dim intRow As Short

        'record selected row in dgv
        If Me.dgvTemplates.Rows.Count = 0 Then
            intRow = -1
        Else
            intRow = Me.dgvTemplates.CurrentRow.Index
        End If


        Me.dgvTemplates.CommitEdit(DataGridViewDataErrorContexts.Commit)
        Me.dgvTemplateAttributes.CommitEdit(DataGridViewDataErrorContexts.Commit)

        Call FillAuditTrailTemp(tblTemplates)
        Call FillAuditTrailTemp(tblTemplateAttributes)

        If boolGuWuOracle Then
            Try
                ta_tblTemplates.Update(tblTemplates)
            Catch ex As DBConcurrencyException
                'ds2005.TBLTEMPLATES.Merge('ds2005.TBLTEMPLATES, True)
            End Try

            Try
                ta_tblTemplateAttributes.Update(tblTemplateAttributes)
            Catch ex As DBConcurrencyException
                'ds2005.TBLTEMPLATEATTRIBUTES.Merge('ds2005.TBLTEMPLATEATTRIBUTES, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_tblTemplatesAcc.Update(tblTemplates)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLTEMPLATES.Merge('ds2005Acc.TBLTEMPLATES, True)
            End Try

            Try
                ta_tblTemplateAttributesAcc.Update(tblTemplateAttributes)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLTEMPLATEATTRIBUTES.Merge('ds2005Acc.TBLTEMPLATEATTRIBUTES, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_tblTemplatesSQLServer.Update(tblTemplates)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLTEMPLATES.Merge('ds2005Acc.TBLTEMPLATES, True)
            End Try

            Try
                ta_tblTemplateAttributesSQLServer.Update(tblTemplateAttributes)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLTEMPLATEATTRIBUTES.Merge('ds2005Acc.TBLTEMPLATEATTRIBUTES, True)
            End Try
        End If

        Dim intTry
        intTry = 0

        'ta_tblTemplates.Fill(tblTemplates)
        'ta_tblTemplateAttributes.Fill(tblTemplateAttributes)

        'update stuff
        Call UpdateStudyName()
        'select current row
        If intRow = -1 Then
        Else
            Me.dgvTemplates.CurrentCell = Me.dgvTemplates.Rows(intRow).Cells("charTemplateName")
        End If
        Call TemplatesAttributesConfigure()
        Call UpdateTabNames()

    End Sub

    Sub DoCancelHooksTab()

        tblHooks.RejectChanges()
        Call HooksInitialize()


    End Sub

    Sub DoCancelComplianceTab()

        tblMeaningOfSig.RejectChanges()
        tblReasonForChange.RejectChanges()
        tblConfigCompliance.RejectChanges()

        Call FillCompliance()
        Call ConfigMOFandRFC()

    End Sub

    Sub DoCancelGlobalTab()

        Dim intRow As Short

        intRow = Me.dgvGlobal.CurrentRow.Index

        Me.dgvGlobal.EndEdit(True)

        tblConfiguration.RejectChanges()

        Call GlobalConfigure()

        Me.dgvGlobal.Rows(intRow).Cells("charConfigTitle").Selected = True


    End Sub

    Sub DoCancelPermissionsManager()

        tblPermissions.RejectChanges()

        Me.dgvPermissions.CommitEdit(DataGridViewDataErrorContexts.Commit)



    End Sub

    Sub DoCancelUserAccountTab()

        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short

        If Me.dgvUsers.Rows.Count = 0 Then
            int1 = -1
        Else
            int1 = Me.dgvUsers.CurrentRow.Index
        End If
        If Me.dgvUserAttributes.Rows.Count = 0 Then
            int2 = -1
        Else
            int2 = Me.dgvUserAttributes.CurrentRow.Index
        End If
        'int2 = Me.dgvUserAttributes.CurrentRow.Index


        'finish any edits in the dgvs
        Me.dgvUsers.CommitEdit(DataGridViewDataErrorContexts.Commit)
        '
        Me.dgvUserAttributes.CommitEdit(DataGridViewDataErrorContexts.Commit)

        tblUserAccounts.RejectChanges()
        tblPersonnel.RejectChanges()

        int3 = Me.dgvUsers.Rows.Count
        'If int1 > int3 - 1 Then
        '    Call ShowUsers(0)
        'Else
        '    Call ShowUsers(int1)
        'End If

        'select last dgvuser and dgvuserattribute
        If int3 = 0 Then

        ElseIf int1 > int3 - 1 Then
            'Me.dgvUsers.Rows(0).Cells("charLastName").Selected = True
            Call ShowUsers(0)
            'Call UserAccountConfigure()
            'Call ConfigureUserAccountAttributes()
        Else
            If Me.dgvUserAttributes.Rows.Count = 0 Then
                Call ConfigureUserAccountAttributes(False)
                Call FillUserboolA(Me.dgvUsers)
            Else
                Call UserAccountInitialize()
                'Me.dgvUsers.Rows(int1).Cells("charLastName").Selected = True
                Call ShowUsers(int1)
                Call ConfigureUserAccountAttributes(True)
                'select last dgvUserAttribute
                Dim dv As System.Data.DataView
                dv = Me.dgvUserAttributes.DataSource
                If int2 > dv.Count - 1 Or int2 = -1 Then
                Else
                    Me.dgvUserAttributes.CurrentCell = Me.dgvUserAttributes.Rows(int2).Cells("charUserID")
                    Me.dgvUserAttributes.Rows(int2).Selected = True

                End If
                Call FillUserboolA(Me.dgvUserAttributes)
            End If
        End If

    End Sub

    Sub DoCancelDropdownboxTab()

        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView
        Dim intRow As Short

        dgv1 = Me.dgvDropdownboxTitle
        dgv2 = Me.dgvDropdownboxContents

        'record selected row in dgv1
        'If dgv1.CurrentRow.Index = Nothing Then
        If dgv1.CurrentRow Is Nothing Then
            'select first row
            intRow = 0
        Else
            intRow = dgv1.CurrentRow.Index
        End If

        'finish any edits in dgv
        dgv1.CommitEdit(DataGridViewDataErrorContexts.Commit)
        dgv2.CommitEdit(DataGridViewDataErrorContexts.Commit)

        'reject changes
        tblDropdownBoxName.RejectChanges()
        tblDropdownBoxContent.RejectChanges()

        dgv1.Rows(intRow).Cells("charDropdownName").Selected = True

        Call DropDownsConfigure()


    End Sub

    Sub DoCancelCorporateAddressTab()

        boolCancelAddresses = True
        boolAddAddresses = False

        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView
        dgv1 = Me.dgvNickNames
        dgv2 = Me.dgvCorporateAddresses

        'finish any edits in dgv2
        dgv1.CommitEdit(DataGridViewDataErrorContexts.Commit)
        dgv2.CommitEdit(DataGridViewDataErrorContexts.Commit)

        'reject changes
        tblCorporateNickNames.RejectChanges()
        tblCorporateAddresses.RejectChanges()

        Call CorporateAddressesInitialize()

        boolCancelAddresses = False

        'select first row of dgv
        If dgv1.Rows.Count = 0 Then
        Else
            dgv1.CurrentCell = dgv1.Rows(0).Cells("charNickname")
        End If

        Call CorporateAddressFilter()


    End Sub

    Sub DoCancelReportTemplatesTab()

        tblTemplates.RejectChanges()
        tblTemplateAttributes.RejectChanges()

        'select first row of dgv
        Dim dgv As DataGridView

        dgv = Me.dgvTemplates
        If dgv.Rows.Count = 0 Then
        Else
            dgv.CurrentCell = dgv.Rows(0).Cells("boolA")
        End If

        Call UpdateStudyName()

        Call TemplatesAttributesConfigure()


    End Sub

    Sub LockReportTemplatesTab(ByVal bool)

        Me.dgvTemplates.ReadOnly = bool
        Me.dgvTemplateAttributes.ReadOnly = bool



        Try
            Me.dgvTemplateAttributes.Columns("TabName").ReadOnly = True
        Catch ex As Exception

        End Try
        Me.cmdResetDefineReports.Enabled = Not (bool)
        Me.cmdAddTemplate.Enabled = Not (bool)


    End Sub

    Sub LockComplianceTab(ByVal bool)

        'Me.pan7.Enabled = Not (bool)

        Me.gbAuditTrail.Enabled = Not (bool)
        Me.cmdAddMOS.Enabled = Not (bool)
        Me.cmdRemoveMOS.Enabled = Not (bool)
        Me.cmdAddRFC.Enabled = Not (bool)
        Me.cmdRemoveRFC.Enabled = Not (bool)


        Try

            Dim dv1 As System.Data.DataView = Me.dgvMOS.DataSource
            dv1.AllowDelete = False
            dv1.AllowNew = False
            dv1.AllowEdit = Not (bool)

            Dim dv2 As System.Data.DataView = Me.dgvRFC.DataSource
            dv2.AllowDelete = False
            dv2.AllowNew = False
            dv2.AllowEdit = Not (bool)

        Catch ex As Exception

        End Try

        'If bool Then

        '    Try

        '        Dim dv1 As system.data.dataview = Me.dgvMOS.DataSource
        '        dv1.AllowDelete = False
        '        dv1.AllowNew = False

        '        Dim dv2 As system.data.dataview = Me.dgvRFC.DataSource
        '        dv2.AllowDelete = False
        '        dv2.AllowNew = False

        '    Catch ex As Exception

        '    End Try

        'Else


        'End If


        'Me.dgvRFC.ReadOnly = bool
        'Me.dgvMOS.ReadOnly = bool

        'Me.cmdResetDefineReports.Enabled = Not (bool)
        'Me.cmdAddTemplate.Enabled = Not (bool)


    End Sub

    Sub LockCorporateAddressTab(ByVal bool)

        Me.dgvNickNames.ReadOnly = bool
        Me.dgvCorporateAddresses.ReadOnly = bool

        Me.cmdAddCorporateAddress.Enabled = Not (bool)
        Me.cmdResetCorporateAddressses.Enabled = Not (bool)


        Try
            Me.dgvCorporateAddresses.Columns("charAddressLabel").ReadOnly = True
        Catch ex As Exception

        End Try

    End Sub

    Sub LockDropdownboxTab(ByVal bool)

        Me.cmdAddDropdownbox.Enabled = Not (bool)
        Me.cmdResetDropdownbox.Enabled = Not (bool)
        Me.cmdOrderDropdownbox.Enabled = Not (bool)

        Me.dgvDropdownboxTitle.ReadOnly = bool
        Me.dgvDropdownboxContents.ReadOnly = bool

    End Sub

    Sub LockCustomFieldCodeTab(ByVal bool)

        Dim var1
        Dim boolF As Boolean

        boolF = boolFormLoad
        boolFormLoad = True

        Me.panFCcmd.Enabled = Not (bool)

        Me.dgvFC.ReadOnly = bool

        boolFormLoad = boolF

    End Sub

    Sub LockPermmissionsTab(ByVal bool)

        Dim var1
        Dim boolF As Boolean

        Me.cmdSelectAllPermissions.Enabled = Not (bool)
        Me.cmdDeselectAllPermissions.Enabled = Not (bool)
        Me.cmdRemovePM.Enabled = Not (bool)
        Me.cmdAddPM.Enabled = Not (bool)

        boolF = boolFormLoad
        boolFormLoad = True
        Me.cbxPermBase.Enabled = Not (bool)
        Me.lvPermissions.Enabled = Not (bool)
        Me.lvPermissionsAdmin.Enabled = Not (bool)
        Me.lvPermissionsFinalReport.Enabled = Not (bool)
        Me.lvPermissionsReportTemplate.Enabled = Not (bool)

        boolFormLoad = boolF

    End Sub

    Sub LockUserAccountTab(ByVal bool)

        Dim var1

        Dim idP As Int64
        Dim intRow As Int32
        Try
            intRow = Me.dgvUsers.CurrentRow.Index
            idP = Me.dgvUsers("ID_TBLPERSONNEL", intRow).Value
        Catch ex As Exception
            idP = 0
        End Try

        'readonly for dgvUsers and Userattributes is tripping error messages
        Dim boolF As Boolean
        boolF = boolFormLoad
        boolFormLoad = True
        Me.dgvUsers.ReadOnly = bool
        Me.dgvUserAttributes.ReadOnly = bool
        boolFormLoad = boolF

        Me.cmdAddUser.Enabled = Not (bool)
        Me.cmdAddUserID.Enabled = Not (bool)
        Me.cmdResetUserAccounts.Enabled = Not (bool)
        Me.cmdEnterPassword.Enabled = Not (bool)

        Me.gbxPassword.Enabled = Not (bool)
        Me.gbSetPerm.Enabled = Not (bool)
        Me.gbWatsonAccount.Enabled = Not (bool)
        Me.gbWindowsAuth.Enabled = Not (bool)
        Me.gbLDAP.Enabled = Not (bool)

        If idP = 1 Then
            Me.cmdAddUserID.Enabled = False
            Me.gbxPassword.Enabled = False
            Me.gbSetPerm.Enabled = False
        Else
            Me.cmdAddUserID.Enabled = Not (bool)
            Me.gbxPassword.Enabled = Not (bool)
            Me.gbSetPerm.Enabled = Not (bool)
        End If

        Me.dgvUsers.Columns("dtActivated").ReadOnly = True
        Me.dgvUsers.Columns("dtDeactivated").ReadOnly = True
        Me.dgvUserAttributes.Columns("dtActivated").ReadOnly = True
        Me.dgvUserAttributes.Columns("dtDeactivated").ReadOnly = True
        Me.dgvUserAttributes.Columns("dtTimeStamp").ReadOnly = True




    End Sub

    Sub LockGlobalParametersAdmin(ByVal bool)
        Try
            Me.dgvGlobal.ReadOnly = bool
            Me.cmdResetGlobal.Enabled = Not (bool)
            Me.cmdBrowseGlobal.Enabled = Not (bool)

            Me.dgvGlobal.Columns("charConfigTitle").ReadOnly = True
            Me.dgvGlobal.Columns("Example").ReadOnly = True

            If bool Then
            Else
                Dim int1 As Short
                Dim str1 As String

                int1 = Me.lbxGlobal.SelectedIndex
                str1 = Me.lbxGlobal.Items(int1)
                Me.dgvGlobal.Columns("charConfigTitle").ReadOnly = True
                If StrComp(str1, "Directory Paths", CompareMethod.Text) = 0 Then
                    Me.dgvGlobal.Columns("charConfigValue").ReadOnly = True
                ElseIf StrComp(str1, "Password Settings", CompareMethod.Text) = 0 Then
                    Me.dgvGlobal.Columns("charConfigValue").ReadOnly = False
                End If

            End If
        Catch ex As Exception

        End Try

    End Sub

    Sub LockPM(bool As Boolean)


        Call EnablepanPM(bool)

    End Sub

    Sub EnablepanPM(bool As Boolean)

        Me.cmdAddPM.Enabled = Not (bool)
        Me.cmdRemovePM.Enabled = Not (bool)
        Me.cbxPermBase.Enabled = Not (bool)
        Me.cmdSelectAllPermissions.Enabled = Not (bool)
        Me.cmdDeselectAllPermissions.Enabled = Not (bool)
        Me.lvPermissionsAdmin.Enabled = Not (bool)
        Me.lvPermissions.Enabled = Not (bool)
        Me.lvPermissionsReportTemplate.Enabled = Not (bool)
        Me.lvPermissionsFinalReport.Enabled = Not (bool)

    End Sub


    Sub LockFCTab(ByVal bool As Boolean)

        'Me.panFC.Enabled = Not (bool)
        Me.panFCcmd.Enabled = Not (bool)

        Dim dv As system.data.dataview
        Dim str2 As String
        Dim dgv As DataGridView
        dgv = Me.dgvFC

        dv = Me.dgvFC.DataSource
        Try
            If bool Then
                dv.AllowDelete = False
                dv.AllowEdit = False
                dv.AllowNew = False
                dgv.ReadOnly = True
            Else
                dv.AllowEdit = True
                dgv.ReadOnly = False
            End If

            str2 = "CHARFIELDCODE"
            dgv.Columns(str2).ReadOnly = bool
            str2 = "CHARDESCRIPTION"
            dgv.Columns(str2).ReadOnly = bool
            str2 = "CHAREXAMPLE"
            dgv.Columns(str2).ReadOnly = bool
        Catch ex As Exception

        End Try


    End Sub

    Sub LockHooksTab(ByVal bool)

        Me.dgvHooks.ReadOnly = bool
        Me.cmdAddHook.Enabled = Not (bool)
        Me.cmdResetHooks.Enabled = Not (bool)
        'Me.cmdRefreshHook.Enabled = Not (bool)

    End Sub

    Sub LockGlobalTab(ByVal bool)

        Try
            Me.dgvGlobal.ReadOnly = bool
            Me.cmdResetGlobal.Enabled = Not (bool)
            Me.cmdBrowseGlobal.Enabled = Not (bool)

            Me.dgvGlobal.Columns("charConfigTitle").ReadOnly = True
            Me.dgvGlobal.Columns("Example").ReadOnly = True

            If bool Then
            Else
                Dim int1 As Short
                Dim str1 As String

                int1 = Me.lbxGlobal.SelectedIndex
                str1 = Me.lbxGlobal.Items(int1)
                Me.dgvGlobal.Columns("charConfigTitle").ReadOnly = True
                If StrComp(str1, "Directory Paths", CompareMethod.Text) = 0 Then
                    Me.dgvGlobal.Columns("charConfigValue").ReadOnly = True
                ElseIf StrComp(str1, "Password Settings", CompareMethod.Text) = 0 Then
                    Me.dgvGlobal.Columns("charConfigValue").ReadOnly = False
                End If

            End If
        Catch ex As Exception

        End Try

    End Sub

    Sub LockAll(ByVal bool As Boolean)

        Call LockGlobalTab(bool)
        Call LockHooksTab(bool)
        Call LockUserAccountTab(bool)
        Call LockDropdownboxTab(bool)
        Call LockCorporateAddressTab(bool)
        Call LockReportTemplatesTab(bool)
        Call LockComplianceTab(bool)
        Call LockFCTab(bool)
        Call LockPM(bool)

    End Sub

    Sub HideAllPages()

        Me.pan1.Visible = False
        Me.pan2.Visible = False
        Me.pan3.Visible = False
        Me.pan4.Visible = False
        Me.pan5.Visible = False
        Me.pan6.Visible = False
        Me.pan7b.Visible = False
        Me.pan8.Visible = False
        Me.panFC.Visible = False


    End Sub

    Sub HidePages(ByVal pan As Panel)

        If boolFormLoad Then
            Exit Sub
        End If

        Dim Count1 As Short
        Dim ctPan As Short
        Dim str1 As String
        Dim strT As String
        Dim var1

        strT = pan.Name 'debug

        ctPan = 9

        Call HideAllPages()

        Dim strF As String
        Dim rows() As DataRow

        strF = "ID_TBLPERMISSIONS = " & id_tblPermissions
        rows = tblPermissions.Select(strF)
        Call EvalLocks(rows)


        Try
            'now make pan visible
            For Count1 = 1 To ctPan

                str1 = ""

                Select Case Count1
                    Case 1, 2, 3, 4, 5, 6, 8
                        str1 = "pan" & Count1
                    Case 7
                        str1 = "pan" & Count1 & "b"
                    Case 9
                        str1 = "panFC"
                End Select


                If StrComp(str1, pan.Name, CompareMethod.Text) = 0 Then

                    'pesky items
                    Select Case Count1 - 1
                        Case 0 ' "User Accounts"
                            'intP = 0
                        Case 1 ' "Dropdownbox Configuration"
                            'intP = 1
                        Case 2 ' "Corporate Addresses"
                            'intP = 2
                            Call UpdateCorpBool()
                        Case 3 ' "Study Template Definitions"
                            'intP = 3
                            Call UpdateStudyName()
                        Case 4 ' "Global Parameters"
                            'intP = 4
                            Call GlobalConfigure()
                        Case 5 ' "Hooks"
                            'intP = 5
                            Call HooksOrder()
                    End Select

                    '20160708 Larry: This next code for some reason is triggering actions in the panel
                    Dim boolF As Boolean = boolFormLoad
                    boolFormLoad = True
                    pan.Visible = True
                    pan.Refresh()
                    boolFormLoad = boolF
                    Exit For
                End If
            Next

            ''now make all other pans invisible
            'For Count1 = 1 To ctPan
            '    'str1 = "pan" & Count1
            '    str1 = ""
            '    Select Case Count1
            '        Case 1, 2, 3, 4, 5, 6, 8
            '            str1 = "pan" & Count1
            '        Case 7
            '            str1 = "pan" & Count1 & "b"
            '        Case 9
            '            str1 = "panFC"
            '    End Select
            '    If StrComp(str1, pan.Name, CompareMethod.Text) = 0 Then
            '    Else
            '        Try
            '            Me.Controls(str1).Visible = False
            '        Catch ex As Exception
            '            'MsgBox(str1 & ":  " & ex.Message)
            '            var1 = ex.Message
            '            var1 = var1
            '        End Try
            '    End If
            'Next

        Catch ex As Exception

        End Try


    End Sub

    Sub lbxTab1Change()

        'If boolFormLoad Then
        '    Exit Sub
        'End If

        Dim int1 As Short
        Dim int2 As Short
        Dim str1 As String
        Dim intP As Short
        Dim tp As Panel

        str1 = Me.lbxTab1.SelectedItem
        If Len(str1) = 0 Then
            Me.lbxTab1.SelectedIndex = 0
            str1 = Me.lbxTab1.SelectedItem
            If Len(str1) = 0 Then
                Exit Sub
            End If
        End If

        If boolFormLoad Then
            Call HideAllPages()
        End If

        'record selected row number
        int1 = Me.lbxTab1.SelectedIndex


        Select Case str1
            Case "User Accounts"
                intP = 0
                tp = Me.pan1
            Case "Dropdownbox Configuration"
                intP = 1
                tp = Me.pan2
            Case "Corporate Addresses"
                intP = 2
                tp = Me.pan3
            Case "Study Template Definitions"
                intP = 3
                tp = Me.pan4
            Case "Global Parameters"
                intP = 4
                tp = Me.pan5
            Case "Hooks"
                intP = 5
                tp = Me.pan6
            Case "Compliance Settings"
                intP = 6
                tp = Me.pan7b
            Case "Custom Field Codes"
                intP = 7
                tp = Me.panFC
            Case "Permissions Manager"
                intP = 8
                tp = Me.pan8
        End Select

        Call HidePages(tp)

        'pesky items

        Select Case str1
            Case "User Accounts"
                intP = 0
            Case "Dropdownbox Configuration"
                intP = 1
            Case "Corporate Addresses"
                intP = 2
                'Call UpdateCorpBool()
            Case "Study Template Definitions"
                intP = 3
                Call UpdateStudyName()
            Case "Global Parameters"
                intP = 4
                Call GlobalConfigure()
                Call SetGlobalParametersControls()
            Case "Hooks"
                intP = 5
                'Call HooksOrder()
        End Select

        'Show(lblRestriction)

        'Dim strF As String
        'Dim rows() As DataRow

        'strF = "ID_TBLPERMISSIONS = " & id_tblPermissions
        'rows = tblPermissions.Select(strF)
        'Call EvalLocks(rows)

        'pesky
        Dim dgv2 As DataGridView
        Try
            dgv2 = Me.dgvFC
            dgv2.Columns(0).Visible = False
        Catch ex As Exception

        End Try

        Try
            dgv2 = Me.dgvDropdownboxContents
            dgv2.Columns("id_tblDropdownBoxContent").Visible = False
            dgv2.Columns("id_tbldropdownboxname").Visible = False

            dgv2.AutoResizeColumns()

            Me.dgvDropdownboxTitle.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        Catch ex As Exception

        End Try

        Try
            dgv2 = Me.dgvCorporateAddresses
            dgv2.Columns(0).Visible = False
            dgv2.AutoResizeColumns()
        Catch ex As Exception

        End Try

        Try
            dgv2 = Me.dgvTemplateAttributes
            dgv2.Columns("ID_TBLTEMPLATES").Visible = False
            dgv2.AutoResizeColumns()
        Catch ex As Exception

        End Try

    End Sub

    Sub DropdownboxInitialize()

        Dim var1
        Dim strF As String
        Dim strS As String
        Dim dtbl1 As System.Data.DataTable
        Dim dtbl2 As System.Data.DataTable
        'Dim dv1 as system.data.dataview
        'Dim dv2 as system.data.dataview
        Dim Count1 As Short
        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView

        dtbl1 = tblDropdownBoxName
        dtbl2 = tblDropdownBoxContent
        dgv1 = Me.dgvDropdownboxTitle
        dgv1.RowHeadersWidth = 25
        dgv1.AllowUserToOrderColumns = False
        dgv1.AllowUserToResizeColumns = True
        dgv1.AllowUserToResizeRows = True
        dgv1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
        dgv1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

        dgv2 = Me.dgvDropdownboxContents
        dgv2.RowHeadersWidth = 25
        dgv2.AllowUserToOrderColumns = False
        dgv2.AllowUserToResizeColumns = True
        dgv2.AllowUserToResizeRows = True
        dgv2.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
        dgv2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
        'dgv2.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

        'initialize dgv1
        'dv1 = dtbl1.DefaultView
        strF = "ID_TBLDROPDOWNBOXNAME > 0"
        strS = "CHARDROPDOWNNAME ASC"
        Dim dv1 As system.data.dataview = New DataView(dtbl1, strF, strS, DataViewRowState.CurrentRows)
        dv1.AllowNew = False
        dv1.AllowDelete = False
        dv1.AllowEdit = False
        dgv1.DataSource = dv1

        dgv1.Columns("id_tbldropdownboxname").Visible = False
        dgv1.Columns("chardropdownname").Visible = True
        dgv1.Columns("chardropdownname").HeaderText = "Dropdown Box Title"
        For Count1 = 0 To dgv1.Columns.Count - 1
            dgv1.Columns(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
            dgv1.Columns(Count1).ReadOnly = True

        Next
        dgv1.Columns("UPSIZE_TS").Visible = False

        dgv1.AutoResizeColumns()
        dgv1.RowHeadersWidth = 25

        'select first row
        dgv1.Rows(0).Cells("chardropdownname").Selected = True
        var1 = dv1(0).Item("id_tbldropdownboxname")

        'initialize dgv2
        strF = "id_tbldropdownboxname = " & var1
        strS = "intOrder ASC, charValue ASC"
        Dim dv2 As system.data.dataview = New DataView(dtbl2, strF, strS, DataViewRowState.CurrentRows)
        dv2.AllowNew = False
        dv2.AllowDelete = False
        dgv2.DataSource = dv2

        For Count1 = 0 To dgv2.Columns.Count - 1
            dgv2.Columns(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        dgv2.Columns("id_tblDropdownBoxContent").Visible = False
        dgv2.Columns("id_tbldropdownboxname").Visible = False
        dgv2.Columns("charValue").Visible = True
        dgv2.Columns("charValue").HeaderText = "Value"
        dgv2.Columns("charAdjective").HeaderText = "Adjective Form"
        dgv2.Columns("charAcronym").HeaderText = "Acronym"
        dgv2.Columns("intOrder").Visible = True
        dgv2.Columns("intOrder").HeaderText = "Display Order"
        dgv2.Columns("intOrder").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv2.Columns("UPSIZE_TS").Visible = False

        dgv2.AutoResizeColumns()
        dgv2.RowHeadersWidth = 25

        Call ShowDropDownColumns(var1)


    End Sub

    Sub DropDownsConfigure()
        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView
        Dim var1
        Dim dv1 As system.data.dataview
        Dim strF As String
        Dim strS As String
        Dim intRow As Short
        Dim dtbl2 As System.Data.DataTable
        Dim int1 As Short
        Dim int2 As Integer

        dgv1 = Me.dgvDropdownboxTitle
        dgv2 = Me.dgvDropdownboxContents
        dv1 = dgv1.DataSource

        'record row
        If dgv1.CurrentRow Is Nothing Then
            'select first row
            'dgv1.Rows(0).Cells("charDropdownName").Selected = True
            intRow = 0
        Else
            intRow = dgv1.CurrentRow.Index
        End If

        var1 = dv1(intRow).Item("id_tbldropdownboxname")

        'initialize dgv2
        dtbl2 = tblDropdownBoxContent
        int2 = dtbl2.Rows.Count

        strF = "id_tbldropdownboxname = " & var1
        strS = "intOrder ASC, charValue ASC"
        Dim dv2 As system.data.dataview = New DataView(dtbl2, strF, strS, DataViewRowState.CurrentRows)
        dv2.AllowNew = False
        dv2.AllowDelete = False
        dgv2.DataSource = dv2

        int1 = dv2.Count 'for debugging

        Call ShowDropDownColumns(var1)



    End Sub

    Sub ShowDropDownColumns(ByVal int1 As Short)
        Dim bool As Boolean
        Dim dgv2 As DataGridView
        Dim var1
        Dim strA As String
        Dim dgv1 As DataGridView

        dgv1 = Me.dgvDropdownboxTitle

        bool = False
        Select Case int1
            Case 3
                bool = True
        End Select

        dgv2 = Me.dgvDropdownboxContents
        dgv2.Columns("charAdjective").Visible = bool
        dgv2.Columns("charAcronym").Visible = bool

        'set cmdOrderReportBody Position
        Dim wd1, wd2
        Dim int2 As Short
        Dim Count1 As Short
        Dim str1 As String
        Dim var2, var3

        int2 = dgv2.Columns.Count
        wd1 = 0
        wd2 = 0
        wd2 = dgv2.RowHeadersWidth
        wd1 = wd1 + wd2
        var3 = 0

        If bool Then 'wrap cells
            dgv2.DefaultCellStyle.WrapMode = DataGridViewTriState.True
            'dgv2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            var2 = dgv2.Width - wd1 - dgv2.Columns("intOrder").Width
            var3 = var2 / 5
            dgv2.Columns("charValue").Width = var3 * 1.95
            dgv2.Columns("charAdjective").Width = var3 * 1.95
            dgv2.Columns("charAcronym").Width = var3
            'wd1 = var2
        Else
            dgv2.DefaultCellStyle.WrapMode = DataGridViewTriState.False
            'dgv2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgv2.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
            For Count1 = 0 To int2 - 1
                If dgv2.Columns(Count1).Visible Then
                    str1 = dgv2.Columns(Count1).Name
                    If StrComp(str1, "intOrder", CompareMethod.Text) = 0 Then
                    Else
                        wd2 = dgv2.Columns(Count1).Width
                        wd1 = wd1 + wd2
                    End If
                End If
            Next
        End If
        var1 = dgv2.Left + wd1
        dgv2.Columns("intOrder").Width = Me.cmdOrderDropdownbox.Width

        If bool Then
            Me.cmdOrderDropdownbox.Left = dgv2.Left + wd1 + (var3 * 1.95 * 2) + var3 'Me.cmdOrderDropdownbox.Width
        Else
            Me.cmdOrderDropdownbox.Left = dgv2.Left + wd1
        End If

        'hide order column
        Dim dv As system.data.dataview
        Dim strS As String
        dv = dgv2.DataSource
        If dgv1.Rows.Count = 0 Then
            GoTo end1
        ElseIf dgv1.CurrentRow Is Nothing Then
            int2 = 0
        Else
            int2 = dgv1.CurrentRow.Index
        End If
        str1 = dgv1.Item("CHARDROPDOWNNAME", int2).Value
        If StrComp(str1, "Anticoagulant/Preservative", CompareMethod.Text) = 0 Then
            dgv2.Columns("intOrder").Visible = False
            strS = "CHARVALUE ASC"
            Me.cmdOrderDropdownbox.Visible = False
        Else
            dgv2.Columns("intOrder").Visible = True
            strS = "INTORDER ASC"
            Me.cmdOrderDropdownbox.Visible = True
        End If
        dv.Sort = strS

        Dim numW
        numW = dgv2.Columns("CHARVALUE").Width
        If numW > 0.85 * dgv2.Width Then
            dgv2.DefaultCellStyle.WrapMode = DataGridViewTriState.True
            'dgv2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders
            dgv2.Columns("CHARVALUE").Width = 0.85 * dgv2.Width

            'dgv2.AllowUserToResizeRows = True
        End If
        dgv2.AutoResizeRows()
end1:

    End Sub

    Sub TemplatesInitialize()

        Dim var1
        Dim strF As String
        Dim strS As String
        Dim dtbl As System.Data.DataTable
        Dim dtbl1 As System.Data.DataTable
        Dim dtblT As System.Data.DataTable
        Dim dvT2 As System.Data.DataView
        Dim dgv As DataGridView
        Dim Count1 As Short
        Dim int1 As Short
        Dim dtblS As System.Data.DataTable
        Dim bool As Boolean

        'do dgvTemplates
        dtbl = tblTemplates
        dtbl1 = tblTemplateAttributes
        dtblT = tblTab1
        dtblS = tblStudies

        'first attempt to remove unbound columns
        'If dtbl.Columns.Contains("boolI") Then
        '    dtbl.Columns.Remove("boolI")
        'End If
        'If dtbl1.Columns.Contains("boolA") Then
        '    dtbl1.Columns.Remove("boolA")
        'End If

        'add unbound bool column to tblTemplates
        bool = dtbl.Columns.Contains("boolA")
        If bool Then
        Else
            Dim col3 As New DataColumn
            col3.DataType = System.Type.GetType("System.Boolean")
            col3.ColumnName = "boolA"
            col3.Caption = "Active"
            col3.AllowDBNull = True
            dtbl.Columns.Add(col3)
        End If

        'add unbound bool column to tblTemplateAttributes
        bool = dtbl1.Columns.Contains("boolI")
        If bool Then
        Else
            Dim col4 As New DataColumn
            col4.DataType = System.Type.GetType("System.Boolean")
            col4.ColumnName = "boolI"
            col4.Caption = "Include"
            col4.AllowDBNull = True
            dtbl1.Columns.Add(col4)
        End If

        dvT2 = dtbl.DefaultView
        dvT2.AllowDelete = False
        dvT2.AllowNew = False
        dvT2.Sort = "charTemplateName ASC"
        dvT2.RowFilter = "boolActive = -1" ' & True
        dvT2.RowStateFilter = DataViewRowState.CurrentRows

        dgv = Me.dgvTemplates
        dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        dgv.RowHeadersWidth = 25
        dgv.AllowUserToOrderColumns = False
        dgv.AllowUserToResizeColumns = True
        dgv.AllowUserToResizeRows = True
        dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)

        dgv.DataSource = dvT2

        'dgv.VirtualMode = True
        'add an unbound column to dgv
        Dim dc1 As New DataGridViewComboBoxColumn
        dc1.Name = "StudyName"
        dc1.DropDownWidth = 160
        dc1.FlatStyle = FlatStyle.Flat
        dc1.Items.Clear()
        dc1.Sorted = True
        'fill box
        For Count1 = 0 To dtblS.Rows.Count - 1
            var1 = dtblS.Rows(Count1).Item("charWatsonStudyName")
            dc1.Items.Add(var1)
        Next

        dgv.Columns.Add(dc1)
        'Note: Values added in needcellvalue event
        dgv.Columns("StudyName").HeaderText = "Study Name"
        dgv.Columns("boolActive").HeaderText = "Active"
        dgv.Columns("boolActive").Visible = False
        dgv.Columns("boolA").HeaderText = "Active"
        dgv.Columns("boolA").Visible = True
        dgv.Columns("charTemplateName").HeaderText = "Template Name"
        dgv.Columns("id_tblStudies").Visible = False
        dgv.Columns("id_tblTemplates").Visible = False
        dgv.Columns("numincr").Visible = False
        dgv.Columns("UPSIZE_TS").Visible = False
        For Count1 = 0 To dgv.Columns.Count - 1
            dgv.Columns(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        're-order cols
        int1 = dgv.Columns.Count
        dgv.Columns("boolA").DisplayIndex = int1 - 1
        dgv.Columns("StudyName").DisplayIndex = int1 - 2

        'select first row
        If dgv.Rows.Count = 0 Then
            var1 = 0
        Else
            dgv.CurrentCell = dgv.Rows(0).Cells("boolA")
            var1 = dgv.Item("id_tblStudies", 0).Value
        End If
        dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)


        'initialize dgvtemplateattributes

        If dgv.Rows.Count = 0 Then
            var1 = 0
        Else
            var1 = dgv.Item("id_tblTemplates", 0).Value
        End If
        strF = "id_tblTemplates = " & var1
        strS = "id_tblTab1 ASC"
        Dim dv As System.Data.DataView
        dv = dtbl1.DefaultView
        dv.RowFilter = strF
        dv.Sort = strS
        dv.RowStateFilter = DataViewRowState.CurrentRows
        dv.AllowNew = False
        dv.AllowDelete = False

        dgv = Me.dgvTemplateAttributes

        dgv.RowHeadersWidth = 25
        dgv.AllowUserToOrderColumns = False
        dgv.AllowUserToResizeColumns = True
        dgv.AllowUserToResizeRows = True
        dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
        dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

        dgv.VirtualMode = True
        dgv.DataSource = dv
        Dim dc As New DataGridViewTextBoxColumn
        dc.Name = "TabName"
        dgv.Columns.Add(dc)

        'make all columns invisible
        For Count1 = 0 To dgv.Columns.Count - 1
            dgv.Columns(Count1).Visible = False
            dgv.Columns(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        dgv.Columns("boolInclude").Visible = False
        dgv.Columns("boolInclude").HeaderText = "Include"
        dgv.Columns("boolI").Visible = True
        dgv.Columns("boolI").HeaderText = "Include"
        dgv.Columns("TabName").HeaderText = "Tab Name"
        dgv.Columns("TabName").Visible = True
        dgv.Columns("numIncr").Visible = False
        dgv.Columns("id_tblTemplates").Visible = False

        dgv.AutoResizeColumns()

        'select first row
        If dv.Count = 0 Then
        Else
            dgv.CurrentCell = dgv.Rows(0).Cells("boolI")
        End If

    End Sub

    Sub TemplatesConfigure()
        'this macro will choose appropriate values in cbxStudyName
        Dim dv As system.data.dataview
        Dim Count1 As Short
        Dim int1 As Short
        Dim str1 As String
        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim var1
        Dim strF As String

        dv = Me.dgvTemplates.DataSource
        int1 = dv.Count
        tbl = tblStudies
        For Count1 = 0 To int1 - 1
            var1 = dv(Count1).Item("id_tblStudies")
            strF = "id_tblStudies = " & var1
            Erase rows
            rows = tbl.Select(strF)
            str1 = rows(0).Item("charWatsonStudyName")
            Me.dgvTemplates.Item("StudyName", Count1).Value = str1
        Next

    End Sub

    Sub TemplatesAttributesConfigure()

        Dim var1
        Dim strF As String
        Dim strS As String
        Dim dtbl As System.Data.DataTable
        Dim dgv As DataGridView
        Dim Count1 As Short
        Dim dtblT As System.Data.DataTable


        Dim intRow As Short
        Dim tbl As System.Data.DataTable
        Dim tblS As System.Data.DataTable
        Dim rows() As DataRow
        Dim var2
        Dim rowsS() As DataRow
        Dim intRows As Short
        Dim int1 As Short

        'check to see if new default tabs have been added
        'intRow = Me.dgvTemplates.CurrentRow.Index
        If Me.dgvTemplates.Rows.Count = 0 Then
            intRow = -1
        ElseIf Me.dgvTemplates.CurrentRow Is Nothing Then
            intRow = 0
            Me.dgvTemplates.Rows(0).Cells("charTemplateName").Selected = True
        Else
            intRow = Me.dgvTemplates.CurrentRow.Index
        End If
        If intRow = -1 Then
            var1 = -1
        Else
            var1 = Me.dgvTemplates.Rows(intRow).Cells("id_tblTemplates").Value
        End If
        'check to see if any other default items have been configured
        tbl = tblTab1
        tblS = tblTemplateAttributes
        strF = "charForm = 'StudyDoc Main' AND boolIncludeinTemplate = -1" ' & True
        strS = "intorder ASC"
        rows = tbl.Select(strF, strS)
        int1 = rows.Length
        For Count1 = 0 To int1 - 1
            var2 = rows(Count1).Item("id_tblTab1")
            strF = "id_tblTemplates = " & var1 & " AND id_tblTab1 = " & var2
            rowsS = tblS.Select(strF)
            intRows = rowsS.Length
            If intRows = 0 Then 'add record
                Dim nr As DataRow = tblS.NewRow
                nr.BeginEdit()
                nr("id_tblTemplates") = var1
                nr("id_tblTab1") = var2
                nr("boolInclude") = -1
                nr("boolI") = True

                nr.EndEdit()
                tblS.Rows.Add(nr)
            End If
        Next

        If boolGuWuOracle Then
            Try
                ta_tblTemplateAttributes.Update(tblTemplateAttributes)
            Catch ex As DBConcurrencyException
                'ds2005.TBLTEMPLATEATTRIBUTES.Merge('ds2005.TBLTEMPLATEATTRIBUTES, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_tblTemplateAttributesAcc.Update(tblTemplateAttributes)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLTEMPLATEATTRIBUTES.Merge('ds2005Acc.TBLTEMPLATEATTRIBUTES, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_tblTemplateAttributesSQLServer.Update(tblTemplateAttributes)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLTEMPLATEATTRIBUTES.Merge('ds2005Acc.TBLTEMPLATEATTRIBUTES, True)
            End Try
        End If

        'initialize dgvtemplateattributes
        dtbl = tblTemplateAttributes
        dtblT = tblTab1
        If Me.dgvTemplates.Rows.Count = 0 Then
            var1 = 0
        Else
            var1 = Me.dgvTemplates.Item("id_tblTemplates", Me.dgvTemplates.CurrentRow.Index).Value
        End If

        strF = "id_tblTemplates = " & var1
        strS = "id_tblTab1 ASC"
        Dim dv As system.data.dataview = New DataView(dtbl, strF, strS, DataViewRowState.CurrentRows)
        'dv = dtbl.DefaultView
        'dv.RowFilter = strF
        'dv.Sort = strS
        'dv.RowStateFilter = DataViewRowState.CurrentRows
        dv.AllowNew = False
        dv.AllowDelete = False
        dgv = Me.dgvTemplateAttributes
        dgv.DataSource = dv

        Call UpdateTabNames()

        dgv.AutoResizeColumns()

        'select first row
        If dv.Count = 0 Then
        Else
            dgv.CurrentCell = dgv.Rows(0).Cells("boolI")
        End If

    End Sub

    Sub ShowTemplate()
        Dim dgv As DataGridView

        If boolFormLoad Then
            Exit Sub
        End If

        If boolAddRow Then
            Exit Sub
        End If

        Dim dv As system.data.dataview

        'dv = tblTemplates.DefaultView
        dv = New DataView(tblTemplates)
        If Me.rbShowActiveTemplates.Checked Then
            dv.RowFilter = "boolActive = -1" ' & True
        ElseIf Me.rbShowInactiveTemplates.Checked Then
            dv.RowFilter = "boolActive = 0" ' & False
        ElseIf Me.rbShowAllTemplates.Checked Then
            dv.RowFilter = ""
        End If
        dv.AllowDelete = False
        dv.AllowNew = False

        dv.Sort = "CHARTEMPLATENAME ASC"

        dgv = Me.dgvTemplates
        dgv.DataSource = dv

        'select first row
        If dgv.Rows.Count = 0 Then
        Else
            dgv.CurrentCell = dgv.Rows(0).Cells("charTemplateName")
        End If

        Call TemplatesAttributesConfigure()

        Call UpdateStudyName()

    End Sub

    Sub HooksConfigure()

        Dim dtbl As System.Data.DataTable
        Dim dgv As DataGridView
        Dim int1 As Short
        Dim Count1 As Short
        Dim str1 As String

        dgv = Me.dgvHooks
        dtbl = tblHooks

        dgv.RowHeadersWidth = 25
        dgv.AllowUserToOrderColumns = False
        dgv.AllowUserToResizeColumns = True
        dgv.AllowUserToResizeRows = True
        dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)

        int1 = dgv.Columns.Count
        For Count1 = 0 To int1 - 1  '
            str1 = dgv.Columns(Count1).Name
            If StrComp(str1, "boolA", CompareMethod.Text) = 0 Then
                dgv.Columns(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            ElseIf StrComp(str1, "boolS", CompareMethod.Text) = 0 Then
                dgv.Columns(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            ElseIf StrComp(str1, "id_tblTab1", CompareMethod.Text) = 0 Then
                dgv.Columns(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            Else
                dgv.Columns(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
            End If
            dgv.Columns(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
            'dgv.Columns(Count1).Visible = False
            dgv.Columns(Count1).DisplayIndex = int1 - 1 'set to high number

        Next

        dgv.Columns("id_tblHooks").Visible = False
        dgv.Columns("BOOLINCLUDE").Visible = False
        dgv.Columns("BOOLSHOW").Visible = False
        dgv.Columns("UPSIZE_TS").Visible = False

        dgv.Columns("boolA").Visible = True
        dgv.Columns("boolA").HeaderText = "Active"
        dgv.Columns("boolA").DisplayIndex = 0

        dgv.Columns("boolS").Visible = True
        dgv.Columns("boolS").HeaderText = "Show"
        dgv.Columns("boolS").DisplayIndex = 1

        dgv.Columns("id_tblTab1").Visible = True
        dgv.Columns("id_tblTab1").HeaderText = "Tab ID"
        dgv.Columns("id_tblTab1").DisplayIndex = 2

        dgv.Columns("CHARHOOK").Visible = True
        dgv.Columns("CHARHOOK").HeaderText = "Hook"
        dgv.Columns("CHARHOOK").DisplayIndex = 3
        dgv.Columns("CHARHOOK").MinimumWidth = 100

        dgv.Columns("CHARUID").Visible = True
        dgv.Columns("CHARUID").HeaderText = "User ID"
        dgv.Columns("CHARUID").DisplayIndex = 4
        dgv.Columns("CHARUID").MinimumWidth = 100

        dgv.Columns("CHARPSWD").Visible = True
        dgv.Columns("CHARPSWD").HeaderText = "Password"
        dgv.Columns("CHARPSWD").DisplayIndex = 5
        dgv.Columns("CHARPSWD").MinimumWidth = 100

        dgv.Columns("CHARCONNECTIONSTRING").Visible = True
        dgv.Columns("CHARCONNECTIONSTRING").HeaderText = "Connection String"
        dgv.Columns("CHARCONNECTIONSTRING").DisplayIndex = 6
        dgv.Columns("CHARCONNECTIONSTRING").MinimumWidth = 250

        dgv.Columns("BOOLERROR").Visible = False


        dgv.AutoResizeColumns()

        If dgv.Rows.Count = 0 Then
        Else
            dgv.CurrentCell = dgv.Rows(0).Cells("CHARHOOK")
        End If

    End Sub

    Sub HooksOrder()
        'pesky
        Dim dgv As DataGridView

        dgv = Me.dgvHooks

        dgv.Columns("boolA").DisplayIndex = 0

        dgv.Columns("boolS").DisplayIndex = 1

        dgv.Columns("id_tblTab1").DisplayIndex = 2

        dgv.Columns("CHARHOOK").DisplayIndex = 3

        dgv.Columns("CHARUID").DisplayIndex = 4

        dgv.Columns("CHARPSWD").DisplayIndex = 5

        dgv.Columns("CHARCONNECTIONSTRING").DisplayIndex = 6
    End Sub

    Sub HooksInitialize()
        Dim dtbl As System.Data.DataTable
        Dim dgv As DataGridView
        Dim bool As Boolean

        dgv = Me.dgvHooks
        dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

        dtbl = tblHooks

        'add unbound column to tblHooks
        'add unbound bool column to tblTemplates
        bool = dtbl.Columns.Contains("boolA")
        If bool Then
        Else
            Dim col3 As New DataColumn
            col3.DataType = System.Type.GetType("System.Boolean")
            col3.ColumnName = "boolA"
            col3.Caption = "Active"
            col3.AllowDBNull = True
            dtbl.Columns.Add(col3)
        End If

        bool = dtbl.Columns.Contains("boolS")
        If bool Then
        Else
            Dim col4 As New DataColumn
            col4.DataType = System.Type.GetType("System.Boolean")
            col4.ColumnName = "boolS"
            col4.Caption = "Show"
            col4.AllowDBNull = True
            dtbl.Columns.Add(col4)
        End If

        'update values in new column
        Call UpdateHookActive()

        'strF = "boolActive = -1" ' & True 'default state is true
        'strS = "charLastName ASC, charFirstName ASC, charMiddleName ASC"
        'Dim dv as system.data.dataview ' = New DataView(dtbl, strF, strS, DataViewRowState.CurrentRows)
        Dim dv As system.data.dataview = New DataView(dtbl)
        dv.RowStateFilter = DataViewRowState.CurrentRows
        dv.AllowNew = False
        dv.AllowDelete = False
        dgv.DataSource = dv
        'select first row

        Call HooksConfigure()


    End Sub

    Sub UserAccountInitialize()

        Dim var1
        Dim strF As String
        Dim strS As String
        Dim dtbl As System.Data.DataTable
        Dim dgv As DataGridView
        Dim bool As Boolean
        Dim Count1 As Short

        dgv = Me.dgvUsers


        'do dgvUsers
        dtbl = tblPersonnel


        If dgv.Columns.Count > 0 Then
            'dgv.Columns("charUserid").Frozen = False
            Try
                'dgv.Columns("charUserid").Frozen = False
                For Count1 = 0 To dgv.Columns.Count - 1
                    dgv.Columns(Count1).Frozen = False
                Next
            Catch ex As Exception

            End Try

        End If

        'add unbound bool column to tblPersonnel
        bool = dtbl.Columns.Contains("boolA")
        If bool Then
        Else
            Dim col3 As New DataColumn
            col3.DataType = System.Type.GetType("System.Boolean")
            col3.ColumnName = "boolA"
            col3.Caption = "A*"
            col3.AllowDBNull = True
            dtbl.Columns.Add(col3)
        End If
        Try
            dgv.Columns("boolA").DisplayIndex = 0
        Catch ex As Exception

        End Try


        strF = "boolActive = -1" ' & True 'default state is true
        strF = "DTACTIVATED IS NOT NULL AND DTDEACTIVATED IS NULL"
        strS = "charLastName ASC, charFirstName ASC, charMiddleName ASC"
        Dim dv As system.data.dataview ' = New DataView(dtbl, strF, strS, DataViewRowState.CurrentRows)
        dv = dtbl.DefaultView
        dv.RowFilter = strF
        dv.Sort = strS
        dv.RowStateFilter = DataViewRowState.CurrentRows
        dv.AllowNew = False
        dv.AllowDelete = False

        dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        dgv.DataSource = dv

        Try
            dgv.Columns("boolA").DisplayIndex = 0
        Catch ex As Exception

        End Try

        Try
            dgv.Columns("boolA").HeaderText = "A*"
        Catch ex As Exception

        End Try

        Call FillUserboolA(dgv) 'fill boolA column

        'select aaAdmin
        Dim str1 As String
        Dim intRow As Int64
        If dv.Count = 0 Then
        Else
            intRow = 0
            'For Count1 = 0 To dv.Count - 1
            '    str1 = NZ(dv(Count1).Item("charLastName"), "aa")
            '    If StrComp(str1, "aaAdmin", CompareMethod.Text) = 0 Then
            '        intRow = Count1
            '        Exit For
            '    End If
            'Next
            dgv.CurrentCell = dgv.Rows(intRow).Cells("charLastName")

            'the following won't get called because of boolformload
            'must run them individually

            Call ConfigureUserAccountAttributes(False)

            Call EvaluateUserAccounts()

            Call FillUserboolA(Me.dgvUserAttributes)

            Call UserIDActions()

        End If




        'do dgvUserAttributes

        If dv.Count = 0 Then
            var1 = 0
        Else
            var1 = dv(0).Item("id_tblPersonnel")
        End If
        dtbl = tblUserAccounts
        strF = "id_tblPersonnel = " & var1 & " AND boolActive = -1" ' & True 'default is true
        strS = "charUserID ASC"
        Dim dv1 As system.data.dataview = New DataView(dtbl, strF, strS, DataViewRowState.CurrentRows)
        dv1.AllowNew = False
        dv1.AllowDelete = False
        dgv = Me.dgvUserAttributes

        If dgv.Columns.Count > 0 Then
            'dgv.Columns("charUserid").Frozen = False
            Try
                'dgv.Columns("charUserid").Frozen = False
                For Count1 = 0 To dgv.Columns.Count - 1
                    dgv.Columns(Count1).Frozen = False
                Next
            Catch ex As Exception

            End Try

        End If

        'add unbound bool column to tblPersonnel
        bool = dtbl.Columns.Contains("boolA")
        If bool Then
        Else
            Dim col4 As New DataColumn
            col4.DataType = System.Type.GetType("System.Boolean")
            col4.ColumnName = "boolA"
            col4.Caption = "A*"
            col4.AllowDBNull = True
            dtbl.Columns.Add(col4)
        End If


        If dgv.Columns.Count > 0 Then
            'dgv.Columns("charUserid").Frozen = False
            Try
                'dgv.Columns("charUserid").Frozen = False
                For Count1 = 0 To dgv.Columns.Count - 1
                    dgv.Columns(Count1).Frozen = False
                Next
            Catch ex As Exception

            End Try
            Try
                dgv.Columns("boolA").DisplayIndex = 0
            Catch ex As Exception

            End Try
            dgv.DataSource = dv1
            'ensure boolA column is first
            dgv.Columns("boolA").DisplayIndex = 0
            Try
                'dgv.Columns("charUserid").Frozen = True
            Catch ex As Exception

            End Try
        Else
            dgv.DataSource = dv1
        End If

        Try
            dgv.Columns("boolA").DisplayIndex = 0
        Catch ex As Exception

        End Try

        Try
            dgv.Columns("boolA").HeaderText = "A*"
        Catch ex As Exception

        End Try


        If dv1.Count = 0 Then
        Else
            dgv.CurrentCell = dgv.Rows(0).Cells("charUserID")
        End If
        dgv.AutoResizeColumns()

        Call FillUserboolA(dgv) 'fill boolA column


    End Sub

    Sub dgvPermissionsConfigure()

        Dim dtbl As DataTable
        Dim var1
        Dim strF As String
        Dim strS As String
        Dim rows() As DataRow
        Dim Count1 As Short
        Dim intCols As Short
        Dim str1 As String
        Dim str2 As String
        Dim boolV As Boolean
        Dim dgv As DataGridView

        Try
            'do dgvPermissions
            dtbl = tblPermissions
            dgv = Me.dgvPermissions

            'strF = "id_tblUserAccounts = " & var1
            strF = "ID_TBLPERMISSIONS > 0"
            'strS = "id_tblPermissions ASC"
            strS = "CHARPERMISSIONSNAME ASC"
            rows = dtbl.Select(strF, strS)

            Dim dv As System.Data.DataView = New DataView(tblPermissions)
            dv.RowFilter = strF
            dv.Sort = strS
            dv.AllowDelete = False
            dv.AllowEdit = False
            dv.AllowNew = False
            'dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            dgv.DataSource = dv


            intCols = dtbl.Columns.Count

            'for some reason, col1 is still visible
            'first make everything invisible
            For Count1 = 0 To intCols - 1
                str1 = dtbl.Columns(Count1).ColumnName
                boolV = False
                dgv.Columns(Count1).Visible = boolV
                dgv.Columns(Count1).HeaderText = str1
            Next


            For Count1 = 0 To intCols - 1
                str1 = dtbl.Columns(Count1).ColumnName
                boolV = False
                Select Case str1
                    Case "CHARPERMISSIONSNAME"
                        boolV = True
                        str2 = "Permissions Group Name"

                End Select

                If boolV Then
                    dgv.Columns(Count1).Visible = boolV
                    dgv.Columns(Count1).HeaderText = str2
                End If

            Next

            dgv.RowHeadersWidth = 25
            With dgv.ColumnHeadersDefaultCellStyle
                .Font = New Font(dgv.Font, FontStyle.Bold)
            End With
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            dgv.AutoResizeColumns()

            'select first row
            dgv.Rows(0).Selected = True

        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

    End Sub

    Sub UserAccountConfigure()

        Dim dgv As DataGridView
        Dim Count1 As Short
        Dim Count2 As Short
        Dim int1 As Short
        Dim str1 As String
        Dim str3 As String
        Dim strCUT As String

        strCUT = "MMM dd, yyyy  HH:mm:ss tt"

        dgv = Me.dgvUserAttributes
        For Count2 = 1 To 2
            Select Case Count2
                Case 1
                    dgv = Me.dgvUsers
                Case 2
                    dgv = Me.dgvUserAttributes
            End Select


            dgv.RowHeadersWidth = 25
            dgv.AllowUserToOrderColumns = False
            dgv.AllowUserToResizeColumns = True
            dgv.AllowUserToResizeRows = True
            dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)


            int1 = dgv.Columns.Count
            For Count1 = 0 To int1 - 1  '
                If Count2 = 3 Then
                    dgv.Columns(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
                Else
                    dgv.Columns(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
                End If
                dgv.Columns(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
                'dgv.Columns(Count1).Visible = False
            Next

        Next


        'configure dgvUsers
        dgv = Me.dgvUsers
        dgv.Columns("id_tblPersonnel").Visible = False

        dgv.Columns("charLastName").Visible = True
        dgv.Columns("charLastName").HeaderText = "Last Name"

        dgv.Columns("charFirstName").Visible = True
        dgv.Columns("charFirstName").HeaderText = "First Name"

        dgv.Columns("charMiddleName").Visible = True
        dgv.Columns("charMiddleName").HeaderText = "Middle Name"

        'dgv.Columns("boolActive").Visible = False
        'dgv.Columns("boolActive").HeaderText = "Active"
        'dgv.Columns("boolActive").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        dgv.Columns("boolActive").Visible = False 'False
        dgv.Columns("boolA").Visible = True 'False
        dgv.Columns("boolA").HeaderText = "A*"
        dgv.Columns("boolActive").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        dgv.Columns("dtActivated").Visible = True
        dgv.Columns("dtActivated").HeaderText = "Date Activated"
        dgv.Columns("dtActivated").DefaultCellStyle.Format = strCUT

        dgv.Columns("dtDeactivated").Visible = True
        dgv.Columns("dtDeactivated").HeaderText = "Date Deactivated"
        dgv.Columns("dtDeactivated").DefaultCellStyle.Format = strCUT

        dgv.Columns("charComments").Visible = True
        dgv.Columns("charComments").HeaderText = "Comments"

        dgv.Columns("CHAREMAILADDRESS").Visible = True
        dgv.Columns("CHAREMAILADDRESS").HeaderText = "eMail Address"

        dgv.Columns("UPSIZE_TS").Visible = False


        Me.dgvUsers.AutoResizeColumns()
        Try
            'dgv.Columns("charFirstName").Frozen = True
        Catch ex As Exception

        End Try

        Me.dgvUsers.Columns("dtActivated").ReadOnly = True
        Me.dgvUsers.Columns("dtDeactivated").ReadOnly = True

        'select 2nd row. 1st row is Admin
        Try
            Me.dgvUsers.CurrentCell = Me.dgvUsers.Rows(0).Cells("CHARLASTNAME")
            Me.dgvUsers.Rows(1).Selected = True
        Catch ex As Exception

        End Try
      


        'configure dgvUserAttributes
        dgv = Me.dgvUserAttributes
        If dgv.Rows.Count = 0 Then
            'GoTo end1
        Else
        End If
        dgv.Columns("id_tblUserAccounts").Visible = False
        dgv.Columns("id_tblUserAccounts").HeaderText = "ID_U"
        dgv.Columns("id_tblPersonnel").Visible = False
        dgv.Columns("id_tblPersonnel").HeaderText = "ID_P"

        dgv.Columns("charUserID").Visible = True
        dgv.Columns("charUserID").HeaderText = "User ID"

        dgv.Columns("charPassword").Visible = False
        dgv.Columns("charPassword").HeaderText = "Password"


        dgv.Columns("boolActive").Visible = False
        dgv.Columns("boolActive").HeaderText = "Active"
        dgv.Columns("boolActive").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        dgv.Columns("boolA").Visible = True 'False
        dgv.Columns("boolA").HeaderText = "A*"

        dgv.Columns("dtActivated").Visible = True
        dgv.Columns("dtActivated").HeaderText = "Date Activated"
        dgv.Columns("dtActivated").DefaultCellStyle.Format = strCUT

        dgv.Columns("dtDeactivated").Visible = True
        dgv.Columns("dtDeactivated").HeaderText = "Date Deactivated"
        dgv.Columns("dtDeactivated").DefaultCellStyle.Format = strCUT

        dgv.Columns("dtTimeStamp").Visible = True
        dgv.Columns("dtTimeStamp").HeaderText = "Last Password" & ChrW(13) & "Change Time"
        dgv.Columns("dtTimeStamp").DefaultCellStyle.Format = strCUT

        dgv.Columns("charComments").Visible = True
        dgv.Columns("charComments").HeaderText = "Comments"

        dgv.Columns("DTLOGONTIME").Visible = True
        dgv.Columns("DTLOGONTIME").HeaderText = "Last Logged In"
        dgv.Columns("dtDeactivated").DefaultCellStyle.Format = strCUT


        dgv.Columns("ID_TBLPERMISSIONS").Visible = False

        dgv.Columns("ID_TBLWATSONACCOUNT").Visible = False
        dgv.Columns("ID_TBLWINDOWSAUTH").Visible = False
        dgv.Columns("CHARLDAP").Visible = False


        dgv.Columns("numincr").Visible = False
        dgv.Columns("UPSIZE_TS").Visible = False


        dgv.Columns("dtActivated").ReadOnly = True
        dgv.Columns("dtDeactivated").ReadOnly = True
        dgv.Columns("dtTimeStamp").ReadOnly = True

        For Count1 = 0 To dgv.Columns.Count - 1
            str1 = dgv.Columns(Count1).Name
            If StrComp(str1, "boolActive", CompareMethod.Text) = 0 Or StrComp(str1, "boolA", CompareMethod.Text) = 0 Then
            Else
                If InStr(1, str1, "bool", CompareMethod.Text) > 0 Then
                    dgv.Columns(Count1).Visible = False
                End If
            End If
        Next

        Me.dgvUserAttributes.AutoResizeColumns()
        Try
            'dgv.Columns("charUserID").Frozen = True
        Catch ex As Exception

        End Try

        'now update Password checkboxes
        Call DVToPasswordCheckboxValues()


        'Call EvaluateUserAccounts()
end1:

    End Sub

    Sub FilllvPermissions()

        'Note: This gets called only at formload and cmdCancel

        Dim dgv As DataGridView
        Dim int1 As Short
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim idP As Int64

        dgv = Me.dgvPermissions

        'configure lvPermissions
        Dim tbl1 As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim strS As String

        Dim boolFL As Boolean


        tbl1 = tblDataTableRowTitles
        strF = "charDataTableName = 'tblPermissions' AND boolInclude = -1 AND CHARTABLEREF IS NULL" ' & True
        strS = "intOrder ASC"
        rows = tbl1.Select(strF, strS)

        Me.lvPermissions.Items.Clear()
        int1 = rows.Length

        boolFL = boolFormLoad
        boolFormLoad = True
        For Count1 = 0 To int1 - 1
            str1 = rows(Count1).Item("charTableRefColumnName")
            str3 = rows(Count1).Item("charRowName")
            Me.lvPermissions.Items.Add(str3)
        Next
        boolFormLoad = boolFL

        'now do lvpermissionsadmin
        strF = "charDataTableName = 'tblPermissionsAdmin' AND boolInclude = -1 AND CHARTABLEREF IS NULL" ' & True
        strS = "intOrder ASC"
        rows = tbl1.Select(strF, strS)

        Me.lvPermissionsAdmin.Items.Clear()
        int1 = rows.Length
        For Count1 = 0 To int1 - 1
            str1 = rows(Count1).Item("charTableRefColumnName")
            str3 = rows(Count1).Item("charRowName")
            Me.lvPermissionsAdmin.Items.Add(str3)
        Next



        'now do lvpermissionsreporttemplate
        strF = "charDataTableName = 'tblPermissionsReportTemplate' AND boolInclude = -1 AND CHARTABLEREF IS NULL" ' & True
        strS = "intOrder ASC"
        rows = tbl1.Select(strF, strS)

        Me.lvPermissionsReportTemplate.Items.Clear()
        int1 = rows.Length
        For Count1 = 0 To int1 - 1
            str1 = rows(Count1).Item("charTableRefColumnName")
            str3 = rows(Count1).Item("charRowName")
            Me.lvPermissionsReportTemplate.Items.Add(str3)
        Next


        'now do lvpermissionsfinalreport
        strF = "charDataTableName = 'tblPermissionsFinalReport' AND boolInclude = -1 AND CHARTABLEREF IS NULL" ' & True
        strS = "intOrder ASC"
        rows = tbl1.Select(strF, strS)

        Me.lvPermissionsFinalReport.Items.Clear()
        int1 = rows.Length
        For Count1 = 0 To int1 - 1
            str1 = rows(Count1).Item("charTableRefColumnName")
            str3 = rows(Count1).Item("charRowName")
            Me.lvPermissionsFinalReport.Items.Add(str3)
        Next



        'the first row is selected in dgvPermissions
        idP = dgv("ID_TBLPERMISSIONS", 0).Value

        boolStopItemCheck = True
        Call ConfigureLVPermissions(idP)
        boolStopItemCheck = False



    End Sub


    Sub ConfigureLVPermissions(idP As Int64)

        Dim Count1 As Short
        Dim str1 As String
        Dim dgv As DataGridView
        Dim dv As System.Data.DataView
        Dim boolA As Boolean
        Dim var1
        Dim strM As String
        Dim strP As String

        dgv = Me.dgvPermissions
        dv = dgv.DataSource
        'Me.lvPermissions.Enabled = True

        'configure lvPermissions
        Dim tbl1 As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim strS As String
        Dim int1 As Short

        Dim rowsPerm() As DataRow
        Dim idUA As Int64

        idUA = idP
        'strF = "ID_TBLUSERACCOUNTS = " & idUA
        strF = "ID_TBLPERMISSIONS = " & idUA
        rowsPerm = tblPermissions.Select(strF)

        strP = "tblPermissions"
        tbl1 = tblDataTableRowTitles
        strF = "charDataTableName = '" & strP & "' AND boolInclude = -1 AND CHARTABLEREF IS NULL" ' & True
        strS = "intOrder ASC"
        rows = tbl1.Select(strF, strS)

        int1 = rows.Length
        Dim boolF As Boolean
        boolF = boolFormLoad
        boolFormLoad = True

        Try
            For Count1 = 0 To int1 - 1
                str1 = rows(Count1).Item("charTableRefColumnName")
                'str3 = rows(Count1).Item("charRowName")
                If dv.Count = 0 Then
                    boolA = False
                Else
                    var1 = rowsPerm(0).Item(str1) 'dv(0).Item(str1)
                    If IsDBNull(var1) Then
                        boolA = False
                    Else
                        boolA = rowsPerm(0).Item(str1) ' dv(0).Item(str1)
                    End If
                End If

                Me.lvPermissions.Items(Count1).Checked = boolA
            Next

            boolFormLoad = boolF

            'If dv.Count = 0 Then
            '    Me.lvPermissions.Enabled = False
            'Else
            '    Me.lvPermissions.Enabled = True
            'End If

            'now do lvpermissionsadmin
            strP = "tblPermissionsAdmin"
            strF = "charDataTableName = '" & strP & "' AND boolInclude = -1 AND CHARTABLEREF IS NULL" ' & True
            strS = "intOrder ASC"
            rows = tbl1.Select(strF, strS)

            int1 = rows.Length

            boolF = boolFormLoad
            boolFormLoad = True
            For Count1 = 0 To int1 - 1
                str1 = rows(Count1).Item("charTableRefColumnName")
                'str3 = rows(Count1).Item("charRowName")
                If dv.Count = 0 Then
                    boolA = False
                Else
                    var1 = rowsPerm(0).Item(str1) ' dv(0).Item(str1)
                    If IsDBNull(var1) Then
                        boolA = False
                    Else
                        boolA = rowsPerm(0).Item(str1) ' dv(0).Item(str1)
                    End If
                End If
                Me.lvPermissionsAdmin.Items(Count1).Checked = boolA
            Next

            boolFormLoad = boolF


            'now do lvpermissionsreporttemplate
            strP = "tblPermissionsReportTemplate"
            strF = "charDataTableName = '" & strP & "' AND boolInclude = -1 AND CHARTABLEREF IS NULL" ' & True
            strS = "intOrder ASC"
            rows = tbl1.Select(strF, strS)

            int1 = rows.Length

            boolF = boolFormLoad
            boolFormLoad = True
            For Count1 = 0 To int1 - 1
                str1 = rows(Count1).Item("charTableRefColumnName")
                'str3 = rows(Count1).Item("charRowName")
                If dv.Count = 0 Then
                    boolA = False
                Else
                    var1 = rowsPerm(0).Item(str1) ' dv(0).Item(str1)
                    If IsDBNull(var1) Then
                        boolA = False
                    Else
                        boolA = rowsPerm(0).Item(str1) ' dv(0).Item(str1)
                    End If
                End If
                Me.lvPermissionsReportTemplate.Items(Count1).Checked = boolA
            Next
            boolFormLoad = boolF



            'now do lvpermissionsfinalreport
            strP = "tblPermissionsFinalReport"
            strF = "charDataTableName = '" & strP & "' AND boolInclude = -1 AND CHARTABLEREF IS NULL" ' & True
            strS = "intOrder ASC"
            rows = tbl1.Select(strF, strS)

            int1 = rows.Length

            boolF = boolFormLoad
            boolFormLoad = True
            For Count1 = 0 To int1 - 1
                str1 = rows(Count1).Item("charTableRefColumnName")
                'str3 = rows(Count1).Item("charRowName")
                If dv.Count = 0 Then
                    boolA = False
                Else
                    var1 = rowsPerm(0).Item(str1) ' dv(0).Item(str1)
                    If IsDBNull(var1) Then
                        boolA = False
                    Else
                        boolA = rowsPerm(0).Item(str1) ' dv(0).Item(str1)
                    End If
                End If
                Me.lvPermissionsFinalReport.Items(Count1).Checked = boolA
            Next
            boolFormLoad = boolF


            'If dv.Count = 0 Then
            '    Me.lvPermissionsAdmin.Enabled = False
            'Else
            '    Me.lvPermissionsAdmin.Enabled = True
            'End If

        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

        boolFormLoad = boolF


    End Sub

    Sub Filllbx1()

        Dim lbx As ListBox = Me.lbx1

        lbx.Items.Clear()

        lbx1.Items.Add("Report Writer")
        lbx1.Items.Add("Report Templates")
        lbx1.Items.Add("Final Reports")
        lbx1.Items.Add("StudyDoc Administration")

        lbx1.SelectedIndex = 0

        Call ShowlvPermissions()

    End Sub

    Sub ShowlvPermissions()

        Dim strM As String
        Dim str1 As String
        Dim boolDo As Boolean
        Dim bool As Boolean

        strM = Me.lbx1.SelectedItem

        boolPermLoad = True

        boolDo = True
        Select Case strM
            Case "StudyDoc Administration"
                Me.lvPermissionsAdmin.Visible = True
                Me.lvPermissions.Visible = False
                Me.lvPermissionsReportTemplate.Visible = False
                Me.lvPermissionsFinalReport.Visible = False
            Case "Report Writer"
                Me.lvPermissions.Visible = True
                Me.lvPermissionsAdmin.Visible = False
                Me.lvPermissionsReportTemplate.Visible = False
                Me.lvPermissionsFinalReport.Visible = False
            Case "Report Templates"
                Me.lvPermissions.Visible = False
                Me.lvPermissionsAdmin.Visible = False
                Me.lvPermissionsReportTemplate.Visible = True
                Me.lvPermissionsFinalReport.Visible = False
            Case "Final Reports"
                Me.lvPermissions.Visible = False
                Me.lvPermissionsAdmin.Visible = False
                Me.lvPermissionsReportTemplate.Visible = False
                Me.lvPermissionsFinalReport.Visible = True
        End Select

        boolPermLoad = False

    End Sub

    Sub ConfigureUserAccountAttributes(boolFromSave As Boolean)

        Dim dv As System.Data.DataView
        Dim dv1 As System.Data.DataView
        Dim intRow As Short
        Dim var1
        Dim strS As String

        'record current UserAccount Row
        'use this if call comes from Save event
        Dim intRowUA As Short
        If Me.dgvUserAttributes.RowCount = 0 Then
            intRowUA = -1
        ElseIf Me.dgvUserAttributes.CurrentRow Is Nothing Then
            intRowUA = 0
        Else
            intRowUA = Me.dgvUserAttributes.CurrentRow.Index
        End If

        'intRow = Me.dgvUsers.CurrentRow.Index

        If Me.dgvUsers.RowCount = 0 Then
            intRow = -1
            var1 = -1
        ElseIf Me.dgvUsers.CurrentRow Is Nothing Then
            intRow = 0
            var1 = Me.dgvUsers.Rows(intRow).Cells("id_tblPersonnel").Value
        Else
            intRow = Me.dgvUsers.CurrentRow.Index
            var1 = Me.dgvUsers.Rows(intRow).Cells("id_tblPersonnel").Value
        End If

        If Len(var1) = 0 Then
            intRow = -1
            var1 = -1
        End If

        'dv = tblUserAccounts.DefaultView
        dv = New DataView(tblUserAccounts)
        Dim boolF As Boolean = boolFormLoad

        boolFormLoad = True
        If Me.rbShowActiveUserIDs.Checked Then
            dv.RowFilter = "id_tblPersonnel = " & var1 & " AND boolActive <> 0" ' & True
        ElseIf Me.rbShowInactiveUserIDs.Checked Then
            dv.RowFilter = "id_tblPersonnel = " & var1 & " AND boolActive = 0" ' & False
        ElseIf Me.rbShowAllUserIDs.Checked Then
            dv.RowFilter = "id_tblPersonnel = " & var1
        End If
        boolFormLoad = boolF
        strS = "charUserID ASC"
        dv.Sort = strS
        dv.AllowNew = False
        dv.AllowDelete = False
        'Me.dgvUserAttributes.DataSource = dv
        Try
            Me.dgvUserAttributes.Columns("charUserID").Frozen = False
        Catch ex As Exception

        End Try

        Dim boolHT As Boolean
        boolHT = boolHold
        boolHold = True
        Me.dgvUserAttributes.DataSource = dv
        boolHold = boolHT
        'Me.dgvUserAttributes.Columns("charUserID").Frozen = True
        Try
            'Me.dgvUserAttributes.Columns("charUserID").Frozen = True
        Catch ex As Exception

        End Try

        'select current row
        If boolFromSave Then
        Else
            intRowUA = 0
        End If

        If dv.Count = 0 Then
            var1 = 0
        Else
            If intRowUA = -1 Then
                intRowUA = 0
            End If
            'Me.dgvUserAttributes.CurrentCell = Me.dgvUserAttributes.Rows(0).Cells("charUserID")
            Try
                Me.dgvUserAttributes.CurrentCell = Me.dgvUserAttributes.Rows(intRowUA).Cells("charUserID")
            Catch ex As Exception

            End Try
            'var1 = dv(0).Item("id_tblUserAccounts")
            'var1 = dv(intRowUA).Item("ID_TBLPERMISSIONS")
            Try
                var1 = dv(intRowUA).Item("ID_TBLPERMISSIONS")
            Catch ex As Exception

            End Try
        End If
        Me.dgvUserAttributes.AutoResizeColumns()

        Dim strF As String
        strF = "ID_TBLPERMISSIONS = " & var1
        Dim rows() As DataRow = tblPermissions.Select(strF)
        Dim str1 As String
        If rows.Length = 0 Then
            boolHT = boolHold
            Try
                boolHold = True
                Me.cbxPermissionsGroup.SelectedIndex = -1
                boolHold = boolHT
            Catch ex As Exception
                boolHold = boolHT
                var1 = ex.Message
                var1 = var1
            End Try

        Else
            str1 = rows(0).Item("CHARPERMISSIONSNAME")
            Try
                boolHT = boolHold
                boolHold = True
                Me.cbxPermissionsGroup.SelectedItem = str1
                'cboCountry.SelectedIndex = cboCountry.FindString(row.Item("CCCOUNTRY").ToString)
                Me.cbxPermissionsGroup.SelectedIndex = Me.cbxPermissionsGroup.FindString(str1)
                boolHold = boolHT
            Catch ex As Exception
                boolHold = boolHT
                var1 = ex.Message
                var1 = var1
            End Try
        End If

        'now update Password checkboxes
        Call DVToPasswordCheckboxValues()


    End Sub


    Sub ShowUsers(ByVal intRow)

        Dim strS As String
        Dim int1 As Short

        If boolFormLoad Then
            Exit Sub
        End If

        Dim dv As System.Data.DataView

        dv = tblPersonnel.DefaultView
        If Me.rbShowActiveUsers.Checked Then
            dv.RowFilter = "boolActive = -1 or boolActive = 1" ' & True
        ElseIf Me.rbShowInactiveUsers.Checked Then
            dv.RowFilter = "boolActive = 0" ' & False
        ElseIf Me.rbShowAllUsers.Checked Then
            dv.RowFilter = ""
        End If

        Me.dgvUsers.Columns("charLastName").Frozen = False
        strS = "charLastName ASC, charFirstName ASC, charMiddleName ASC"
        dv.Sort = strS
        dv.AllowDelete = False
        dv.AllowNew = False

        Me.dgvUsers.DataSource = dv
        Me.dgvUsers.AutoResizeColumns()
        Try
            'Me.dgvUsers.Columns("charLastName").Frozen = True
        Catch ex As Exception

        End Try

        'select row
        If dv.Count = 0 Then
        Else
            If intRow > Me.dgvUsers.Rows.Count - 1 Or intRow = -1 Then
                intRow = 0
            Else
            End If
            Me.dgvUsers.CurrentCell = Me.dgvUsers.Rows(intRow).Cells("charLastName")
        End If

        If Me.dgvUserAttributes.Rows.Count = 0 Then
            int1 = -1
        ElseIf Me.dgvUserAttributes.CurrentRow Is Nothing Then
            int1 = 0
        Else
            int1 = Me.dgvUserAttributes.CurrentRow.Index
        End If

        Call FillUserboolA(Me.dgvUsers) 'fill boolA column
        Call FillUserboolA(Me.dgvUserAttributes) 'fill boolA column

        Call ShowAccounts(int1)

    End Sub

    Sub ShowAccounts(ByVal introw)

        Dim dv As System.Data.DataView
        Dim dv1 As System.Data.DataView
        Dim var1
        Dim intRowU As Short
        Dim strS As String
        Dim strF As String

        If boolFormLoad Then
            Exit Sub
        End If

        If Me.dgvUsers.Rows.Count = 0 Then
            var1 = 0
        ElseIf Me.dgvUsers.CurrentRow Is Nothing Then
            var1 = 0
        Else
            intRowU = Me.dgvUsers.CurrentRow.Index
            var1 = Me.dgvUsers.Item("id_tblPersonnel", intRowU).Value
        End If

        If Len(var1) = 0 Then
            var1 = 0
        End If

        dv = tblUserAccounts.DefaultView

        boolHold = True



        'If Me.rbShowActiveUserIDs.Checked Then
        '    dv.RowFilter = "id_tblPersonnel = " & var1 & "AND boolActive = -1" ' & True
        'ElseIf Me.rbShowInactiveUserIDs.Checked Then
        '    dv.RowFilter = "id_tblPersonnel = " & var1 & "AND boolActive = 0" ' & False
        'ElseIf Me.rbShowAllUserIDs.Checked Then
        '    dv.RowFilter = "id_tblPersonnel = " & var1
        'End If

        If Me.rbShowActiveUserIDs.Checked Then
            strF = "DTACTIVATED IS NOT NULL AND DTDEACTIVATED IS NULL"
            dv.RowFilter = "id_tblPersonnel = " & var1 & "AND " & strF
        ElseIf Me.rbShowInactiveUserIDs.Checked Then
            strF = "DTACTIVATED IS NOT NULL AND DTDEACTIVATED IS NOT NULL"
            dv.RowFilter = "id_tblPersonnel = " & var1 & "AND " & strF
        ElseIf Me.rbShowAllUserIDs.Checked Then
            dv.RowFilter = "id_tblPersonnel = " & var1
        End If

        Me.dgvUserAttributes.Columns("charUserID").Frozen = False
        strS = "charUserID ASC"
        dv.Sort = strS
        dv.AllowDelete = False
        dv.AllowNew = False

        boolHold = True
        Me.dgvUserAttributes.DataSource = dv
        Me.dgvUserAttributes.AutoResizeColumns()
        Try
            'Me.dgvUserAttributes.Columns("charUserID").Frozen = True
        Catch ex As Exception

        End Try


        'select first row
        If dv.Count = 0 Then
            var1 = 0
        Else
            If introw > dv.Count - 1 Then
                introw = 0
            End If
            Me.dgvUserAttributes.CurrentCell = Me.dgvUserAttributes.Rows(introw).Cells("charUserID")
            introw = Me.dgvUserAttributes.CurrentRow.Index
            'var1 = Me.dgvUserAttributes.Item("id_tblUserAccounts", introw).Value
            var1 = Me.dgvUserAttributes.Item("ID_TBLPERMISSIONS", introw).Value
        End If
        boolHold = False

        Call FillUserboolA(Me.dgvUserAttributes)

    End Sub

    Sub ConfigureAccountPermissionsGroup()

        Dim dv As System.Data.DataView
        Dim intRow As Short
        Dim var1
        Dim dgv As DataGridView
        Dim str1 As String
        Dim idP As Int64
        Dim rows() As DataRow
        Dim cbx As ComboBox = Me.cbxPermissionsGroup

        If boolFormLoad Or boolAddAcct Or boolAddUser Then
            Exit Sub
        End If

        dgv = Me.dgvUserAttributes
        intRow = dgv.CurrentRow.Index
        dv = dgv.DataSource
        'get value
        idP = dv(intRow).Item("ID_TBLPERMISSIONS")
        rows = tblPermissions.Select("ID_TBLPERMISSIONS = " & idP)
        str1 = rows(0).Item("CHARPERMISSIONSNAME")

        Try
            'Note: displayvalue = id_tblpermissions
            'seletecteditem must be id_tblpermissions
            cbx.SelectedItem = str1

            cbx.SelectedIndex = cbx.FindStringExact(str1)

        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try



    End Sub

    Function PasswordValidate() As Boolean

        'Don't need this anymore
        'UserID form forces a password

        PasswordValidate = True

        Exit Function


        If boolAddUserAccount Then
            PasswordValidate = True
            Exit Function
        End If

        PasswordValidate = False

        Dim dgv As DataGridView = Me.dgvUserAttributes

        PasswordValidate = True

        Dim int1 As Int32
        Dim Count1 As Int32
        Dim intRows As Int32
        Dim var1
        Dim strM As String
        Dim str1 As String


        intRows = dgv.Rows.Count
        For Count1 = 0 To intRows - 1
            var1 = dgv("CHARPASSWORD", Count1).Value
            If IsDBNull(var1) Then
                str1 = NZ(dgv("CHARUSERID", Count1).Value, "New UserID")
                strM = "Password for UserID: '" & str1 & "' has not been set."
                MsgBox(strM, vbInformation, "Invalid action...")
                GoTo end1
            Else
                If Len(var1) = 0 Then
                    str1 = NZ(dgv("CHARUSERID", Count1).Value, "New UserID")
                    strM = "Password for UserID: '" & str1 & "' has not been set."
                    MsgBox(strM, vbInformation, "Invalid action...")
                    GoTo end1
                End If
            End If

        Next

        PasswordValidate = True

end1:


    End Function

    Function UserAccountValidate() As Boolean

        Dim intRow As Short
        Dim intCol As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim var1, var3, var4, var5
        Dim col As DataGridViewColumn
        Dim dgv As DataGridView
        Dim dtbl As System.Data.DataTable
        Dim ct1 As Short
        Dim Count1 As Short
        Dim strMsg As String
        Dim int2 As Int64

        UserAccountValidate = True
        If boolFormLoad Or boolAddAcct Or boolAddUser Then
            Exit Function
        End If

        If Me.cmdEdit.Enabled Then
            Exit Function
        End If

        If boolHold Then
            Exit Function
        End If

        dgv = Me.dgvUserAttributes
        str1 = dgv.Columns(intCol).Name
        intRow = dgv.CurrentRow.Index 'e.RowIndex

        'ensure fields are not blank
        For Each col In dgv.Columns
            If StrComp(col.Name, "charUserID", CompareMethod.Text) = 0 Then
                var1 = Trim(NZ(dgv(col.Index, intRow).Value, " "))
                If Len(NZ(var1, "")) = 0 Then
                    'e.Cancel = True
                    dgv.CurrentCell = dgv.Rows(intRow).Cells(col.Index)
                    If boolAddUserAccount Then
                    Else
                        MsgBox("The UserID field cannot be blank,", MsgBoxStyle.Information, "Invalid entry...")
                        UserAccountValidate = False
                        dgv.BeginEdit(True)
                        GoTo end1
                    End If
                End If
            End If
            If StrComp(col.Name, "charPassword", CompareMethod.Text) = 0 Then
                var1 = NZ(dgv(col.Index, intRow).Value, " ")
                If Len(NZ(var1, "")) = 0 Then
                    'e.Cancel = True
                    dgv.CurrentCell = dgv.Rows(intRow).Cells(col.Index)
                    If boolAddUserAccount Then
                    Else
                        MsgBox("The Password field cannot be blank,", MsgBoxStyle.Information, "Invalid entry...")
                        UserAccountValidate = False
                        dgv.BeginEdit(True)
                        GoTo end1
                    End If
                End If
            End If
        Next

        'now ensure that user id is unique
        dtbl = tblUserAccounts
        ct1 = dtbl.Rows.Count
        var1 = dgv("charUserID", intRow).Value
        var4 = dgv("id_tblUserAccounts", intRow).Value
        str1 = Trim(NZ(var1, ""))
        For Count1 = 0 To ct1 - 1
            var3 = dtbl.Rows(Count1).Item("id_tblUserAccounts")
            If var3 = var4 Then
            Else
                var1 = dtbl.Rows(Count1).Item("charUserID")
                str2 = Trim(NZ(var1, ""))
                If StrComp(str1, str2, CompareMethod.Text) = 0 Then
                    'e.Cancel = True
                    Dim drow() As DataRow
                    int2 = dtbl.Rows(Count1).Item("id_tblPersonnel")
                    drow = tblPersonnel.Select("id_tblPersonnel = " & int2)
                    str3 = drow(0).Item("charFirstName") & " " & drow(0).Item("charLastName")
                    var5 = drow(0).Item("boolActive")
                    If var5 = 0 Then
                        str3 = "inactive user " & str3
                    Else
                        str3 = "active user " & str3
                    End If
                    var3 = dtbl.Rows(Count1).Item("boolActive")
                    If var3 = 0 Then
                        strMsg = "An inactive user id labeled " & str2 & " assigned to " & str3 & " already exists."
                    Else
                        strMsg = "An active user id labeled " & str2 & " assigned to " & str3 & " already exists."
                    End If

                    MsgBox(strMsg, MsgBoxStyle.Information, "Invalid entry...")
                    UserAccountValidate = False
                    dgv.BeginEdit(True)
                    GoTo end1

                End If
            End If
        Next

end1:

    End Function

    Sub UpdateHookActive()

        Dim tbl As System.Data.DataTable
        Dim Count1 As Short
        Dim int1 As Short
        Dim int2 As Short
        Dim bool As Boolean
        Dim bool1 As Boolean

        tbl = tblHooks
        int1 = tbl.Rows.Count
        For Count1 = 0 To int1 - 1
            int2 = NZ(tbl.Rows(Count1).Item("BOOLINCLUDE"), -1)
            If int2 = -1 Then
                bool = True
            Else
                bool = False
            End If
            int2 = NZ(tbl.Rows(Count1).Item("BOOLSHOW"), 0)
            If int2 = -1 Then
                bool1 = True
            Else
                bool1 = False
            End If

            tbl.Rows(Count1).BeginEdit()
            tbl.Rows(Count1).Item("boolA") = bool
            tbl.Rows(Count1).Item("boolS") = bool1
            tbl.Rows(Count1).EndEdit()
        Next

    End Sub

    Sub UpdateTabNames()
        Dim tblS As System.Data.DataTable
        Dim dgv As DataGridView
        Dim intRows As Short
        Dim var1, var2
        Dim dvT2 As system.data.dataview
        Dim row() As DataRow
        Dim strF As String
        Dim Count1 As Short

        dgv = Me.dgvTemplateAttributes
        'intCol = e.ColumnIndex
        intRows = dgv.Rows.Count
        'str1 = dgv.Columns("StudyName").Name
        dvT2 = dgv.DataSource
        tblS = tblTab1

        For Count1 = 0 To intRows - 1
            'fill Study column with values
            var1 = dvT2(Count1).Item("id_tblTab1")
            strF = "id_tblTab1 = " & var1
            row = tblS.Select(strF)
            var2 = row(0).Item("charItem")
            dgv("TabName", Count1).Value = var2

            'enter appropriate bool value
            var1 = dvT2(Count1).Item("boolInclude")
            If var1 = -1 Then
                dvT2(Count1).Item("boolI") = True
            Else
                dvT2(Count1).Item("boolI") = False
            End If
            dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
            dgv.Update()
        Next

    End Sub

    Sub UpdateStudyName()

        Dim tblS As System.Data.DataTable
        Dim dgv As DataGridView
        Dim intRows As Short
        Dim var1, var2
        Dim dvT2 As System.Data.DataView
        Dim row() As DataRow
        Dim strF As String
        Dim Count1 As Short

        Try
            dgv = Me.dgvTemplates
            'intCol = e.ColumnIndex
            intRows = dgv.Rows.Count
            'str1 = dgv.Columns("StudyName").Name
            dvT2 = dgv.DataSource
            For Count1 = 0 To intRows - 1
                'fill Study column with values
                tblS = tblStudies
                var1 = dvT2(Count1).Item("id_tblStudies")
                If Len(NZ(var1, "")) = 0 Then
                Else
                    'find item in tbls
                    Erase row
                    strF = "id_tblStudies = " & var1
                    row = tblS.Select(strF)
                    var2 = row(0).Item("charWatsonStudyName")
                    'e.Value = var2
                    dgv("StudyName", Count1).Value = var2
                End If

                'enter appropriate bool value
                var1 = dvT2(Count1).Item("boolActive")
                If var1 = 0 Then
                    dgv("boolA", Count1).Value = False
                Else
                    dgv("boolA", Count1).Value = True
                End If
                dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
                dgv.Update()

            Next

            dgv.AutoResizeColumns()
        Catch ex As Exception

        End Try


    End Sub


    Sub ShowAddresses()
        'Dim dv as system.data.dataview
        Dim dv1 As system.data.dataview
        Dim var1
        Dim intRowU As Short
        Dim strS As String
        Dim intRow As Short
        Dim tbl As System.Data.DataTable

        If boolFormLoad Then
            Exit Sub
        End If

        If Me.dgvNickNames.Rows.Count = 0 Then
            var1 = 0
        Else
            intRowU = Me.dgvNickNames.CurrentRow.Index
            var1 = Me.dgvNickNames.Item("id_tblCorporateNickNames", intRowU).Value
        End If

        tbl = tblCorporateNickNames
        Dim dv As system.data.dataview = New DataView(tbl)

        If Me.rbShowActiveAddresses.Checked Then
            dv.RowFilter = "charNickName <> '[None]' AND boolInclude = -1" ' & True
        ElseIf Me.rbShowInactiveAddresses.Checked Then
            dv.RowFilter = "charNickName <> '[None]' AND boolInclude = 0" ' & False
        ElseIf Me.rbShowAllAddresses.Checked Then
            dv.RowFilter = "charNickName <> '[None]'"
        End If

        strS = "charNickName ASC"
        dv.Sort = strS
        dv.AllowDelete = False
        dv.AllowNew = False
        Me.dgvNickNames.DataSource = dv
        Me.dgvNickNames.AutoResizeColumns()


        'select first row of dgv
        If dv.Count = 0 Then
            var1 = 0
        Else
            If intRowU > dv.Count - 1 Then
                intRowU = 0
            End If
            Me.dgvNickNames.CurrentCell = Me.dgvNickNames.Rows(intRowU).Cells("charNickName")
            intRow = Me.dgvNickNames.CurrentRow.Index
            var1 = Me.dgvNickNames.Item("id_tblCorporateNickNames", intRow).Value
        End If

        dv1 = tblCorporateAddresses.DefaultView
        dv1.RowFilter = "id_tblCorporateNickNames = " & var1
        dv1.AllowDelete = False
        dv1.AllowNew = False
        Me.dgvCorporateAddresses.DataSource = dv1
        Me.dgvCorporateAddresses.AutoResizeColumns()

    End Sub

    Sub EvaluateUserAccounts()

        Dim dv As system.data.dataview
        Dim bool As Boolean

        Dim dgv As DataGridView
        Dim var1, var2
        Dim int1 As Short

        If Me.cmdEdit.Enabled Then
            bool = False
        Else
            bool = True
        End If

        If Me.chkEditMode.Checked Then
            bool = True
        Else
            bool = False
        End If

        dgv = Me.dgvUsers
        'int1 = dgv.CurrentRow.Index

        If dgv.RowCount = 0 Then
            Me.cmdAddUserID.Enabled = False
            Me.gbxPassword.Enabled = False
            Me.cmdEnterPassword.Enabled = False
            Me.gbSetPerm.Enabled = False
            Me.gbWatsonAccount.Enabled = False
            Me.gbWindowsAuth.Enabled = False

            Exit Sub

        ElseIf dgv.CurrentRow Is Nothing Then
            int1 = 0
        Else
            int1 = dgv.CurrentRow.Index
        End If
        var1 = NZ(dgv("charLastName", int1).Value, "")


        Call SetWA()


    End Sub


    Function GlobalEvaluateUser() As Boolean

        Dim str1 As String
        Dim ctPB As Short
        Dim ctPBMax As Short
        Dim boolA As Short
        Dim boolAllow As Boolean
        Dim boolE As Boolean

        If Me.chkEditMode.Checked Then
            boolE = True
        Else
            boolE = False
        End If


        If boolFormLoad Then
            Exit Function
        End If

        boolAddRow = True

        GlobalEvaluateUser = False

        Cursor.Current = Cursors.WaitCursor

        Dim strF As String
        Dim bool As Boolean
        Dim rows() As DataRow
        strF = "ID_TBLPERMISSIONS = " & id_tblPermissions
        rows = tblPermissions.Select(strF)

        boolAllow = True
        If rows.Length = 0 Then
            boolAllow = False
            GlobalEvaluateUser = False
        Else
        End If


        If boolAllow Then 'continue evaluating
            boolA = BOOLUSERACCOUNTS
            If boolA = 0 Then
                bool = False
            Else
                bool = True
                GlobalEvaluateUser = True
            End If
            If boolE Then
                Call LockUserAccountTab(Not (bool))
            Else
                Call LockAll(True)
            End If
        Else
            Call LockAll(True)
        End If

end1:

    End Function

    Sub DVToPasswordCheckboxValues()

        Dim dgv As DataGridView
        Dim dv As System.Data.DataView
        Dim intRow As Short
        Dim var1

        dgv = Me.dgvUserAttributes
        dv = dgv.DataSource

        boolStopPswdCheck = True

        If dv.Count = 0 Then
            var1 = False
            Me.chkChangePasswordAtNextLogon.Checked = var1
            Me.chkUserCannotChangePassword.Checked = var1
            Me.chkPasswordNeverExpires.Checked = var1
            Me.chkAccountIsLockedOut.Checked = var1

            'Me.gbxPassword.Enabled = False
        Else
            'Me.gbxPassword.Enabled = True
            'intRow = dgv.CurrentRow.Index
            If dgv.RowCount = 0 Then
            Else

                '****

                'Dim iRowIndex As Integer
                'For i As Integer = 0 To dgv.SelectedCells.Count - 1

                '    iRowIndex = dgv.SelectedCells.Item(i).RowIndex
                '    iRowIndex = iRowIndex
                'Next

                '****

                If dgv.CurrentRow Is Nothing Then
                    intRow = 0
                Else
                    intRow = dgv.CurrentRow.Index
                End If

                ''debug
                'var1 = dgv("ID_TBLUSERACCOUNTS", intRow).Value
                'MsgBox(var1)

                var1 = NZ(dv(intRow).Item("boolChangePasswordAtNextLogon"), 0)
                Me.chkChangePasswordAtNextLogon.Checked = var1

                var1 = NZ(dv(intRow).Item("boolUserCannotChangePassword"), 0)
                Me.chkUserCannotChangePassword.Checked = var1

                var1 = NZ(dv(intRow).Item("boolPasswordNeverExpires"), 0)
                Me.chkPasswordNeverExpires.Checked = var1

                var1 = NZ(dv(intRow).Item("boolAccountIsLockedOut"), 0)
                Me.chkAccountIsLockedOut.Checked = var1

            End If

        End If

        boolStopPswdCheck = False

    End Sub

    Sub PasswordCheckboxValuesToDV()

        Dim dgv As DataGridView
        Dim dv As System.Data.DataView
        Dim intRow As Short
        Dim var1, var2

        boolStopPswdCheck = True

        dgv = Me.dgvUserAttributes
        dv = dgv.DataSource
        If dv.Count = 0 Then
        Else
            intRow = dgv.CurrentRow.Index

            dv(intRow).BeginEdit()
            var1 = Me.chkChangePasswordAtNextLogon.Checked
            If var1 Then
                var2 = -1
            Else
                var2 = 0
            End If
            dv(intRow).Item("boolChangePasswordAtNextLogon") = var2

            var1 = Me.chkUserCannotChangePassword.Checked
            If var1 Then
                var2 = -1
            Else
                var2 = 0
            End If
            dv(intRow).Item("boolUserCannotChangePassword") = var2

            var1 = Me.chkPasswordNeverExpires.Checked
            If var1 Then
                var2 = -1
            Else
                var2 = 0
            End If
            dv(intRow).Item("boolPasswordNeverExpires") = var2

            var1 = Me.chkAccountIsLockedOut.Checked
            If var1 Then
                var2 = -1
            Else
                var2 = 0
            End If
            dv(intRow).Item("boolAccountIsLockedOut") = var2

            dv(intRow).EndEdit()
        End If

        boolStopPswdCheck = False

    End Sub

    Sub PasswordCheckboxes(boolPswdEx As Boolean)

        If Me.cmdEdit.Enabled Then
            Exit Sub
        End If

        If boolHold Then
            Exit Sub
        End If

        Dim bool1 As Boolean

        If boolPswdEx Then
            bool1 = Me.chkPasswordNeverExpires.Checked
            If bool1 Then
                Me.chkChangePasswordAtNextLogon.Checked = False
            End If
        Else
            bool1 = Me.chkChangePasswordAtNextLogon.Checked
            If bool1 Then
                Me.chkPasswordNeverExpires.Checked = False
            End If
        End If


    End Sub

    Sub PasswordInitialize()

        Dim tbl1 As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim strS As String
        Dim str1 As String
        Dim int1 As Short
        Dim int2 As Short
        Dim Count1 As Short
        Dim dgv As DataGridView

        'fill lbx
        tbl1 = tblTab1
        strF = "intForm = 5"
        strS = "intOrder ASC"
        rows = tbl1.Select(strF, strS)
        int1 = rows.Length
        Me.lbxGlobal.Items.Clear()
        If int1 = 0 Then
            GoTo end1
        Else
            For Count1 = 0 To int1 - 1
                str1 = rows(Count1).Item("charItem")
                Me.lbxGlobal.Items.Add(str1)
            Next
        End If
        'select first item
        Me.lbxGlobal.SelectedIndex = 0

        'now initialize dgv
        int1 = Me.lbxGlobal.SelectedIndex
        dgv = Me.dgvGlobal
        dgv.RowHeadersWidth = 25
        dgv.AllowUserToOrderColumns = False
        dgv.AllowUserToResizeColumns = True
        dgv.AllowUserToResizeRows = True
        dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
        dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

        tbl1 = tblConfiguration

        str1 = Me.lbxGlobal.Items(int1)
        strF = "charConfigCategory = '" & str1 & "'"
        Dim dv1 As system.data.dataview = New DataView(tbl1, strF, strS, DataViewRowState.CurrentRows)
        dv1.AllowNew = False
        dv1.AllowDelete = False
        dgv.DataSource = dv1

        'configure dgv
        'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        'dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        int1 = dgv.Columns.Count
        For Count1 = 0 To int1 - 1
            dgv.Columns(Count1).Visible = False
            dgv.Columns(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
            'dgv.Columns(Count1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Next
        dgv.Columns("charConfigTitle").Visible = True
        dgv.Columns("charConfigTitle").HeaderText = "Item"
        dgv.Columns("charConfigTitle").ReadOnly = True
        dgv.Columns("charConfigValue").Visible = True
        dgv.Columns("charConfigValue").HeaderText = "Value"
        dgv.Columns("Example").Visible = False

        dgv.AutoResizeColumns()
        dgv.RowHeadersWidth = 25

end1:

    End Sub

    Sub GlobalInitialize()

        Dim tbl1 As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim strS As String
        Dim str1 As String
        Dim int1 As Short
        Dim int2 As Short
        Dim Count1 As Short
        Dim dgv As DataGridView

        'fill lbx
        tbl1 = tblTab1
        strF = "intForm = 3 and ID_TBLTAB1 <> 22"
        rows = tbl1.Select(strF, strS)
        int1 = rows.Length
        Me.lbxGlobal.Items.Clear()
        If int1 = 0 Then
            GoTo end1
        Else
            For Count1 = 0 To int1 - 1
                str1 = rows(Count1).Item("charItem")
                Me.lbxGlobal.Items.Add(str1)
            Next
        End If
        'select first item
        Me.lbxGlobal.SelectedIndex = 0

        'now initialize dgv
        int1 = Me.lbxGlobal.SelectedIndex
        dgv = Me.dgvGlobal
        dgv.RowHeadersWidth = 25
        dgv.AllowUserToOrderColumns = False
        dgv.AllowUserToResizeColumns = True
        dgv.AllowUserToResizeRows = True
        dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
        dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

        tbl1 = tblConfiguration

        str1 = Me.lbxGlobal.Items(int1)
        strF = "charConfigCategory = '" & str1 & "'"

        If StrComp(str1, "Global Settings", CompareMethod.Text) = 0 Then
            strS = "CHARCONFIGTITLE ASC"
        Else
            strS = "INTORDER ASC"
        End If

        Dim dv1 As system.data.dataview = New DataView(tbl1, strF, strS, DataViewRowState.CurrentRows)
        dv1.AllowNew = False
        dv1.AllowDelete = False
        dgv.DataSource = dv1

        'configure dgv
        'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        'dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        int1 = dgv.Columns.Count
        For Count1 = 0 To int1 - 1
            dgv.Columns(Count1).Visible = False
            dgv.Columns(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
            'dgv.Columns(Count1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Next
        dgv.Columns("charConfigTitle").Visible = True
        dgv.Columns("charConfigTitle").HeaderText = "Item"
        dgv.Columns("charConfigTitle").ReadOnly = True
        dgv.Columns("charConfigValue").Visible = True
        dgv.Columns("charConfigValue").HeaderText = "Value"
        dgv.Columns("Example").Visible = False

        dgv.AutoResizeColumns()
        dgv.RowHeadersWidth = 25

end1:

    End Sub

    Sub GlobalConfigure()

        Dim tbl1 As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim strS As String
        Dim str1 As String
        Dim int1 As Short
        Dim int2 As Short
        Dim Count1 As Short
        Dim dgv As DataGridView
        Dim boolDir As Boolean
        Dim boolFile As Boolean
        Dim strL As String
        Dim var1, var2

        int1 = Me.lbxGlobal.SelectedIndex
        dgv = Me.dgvGlobal
        tbl1 = tblConfiguration
        strL = Me.lbxGlobal.Items(int1)
        strF = "charConfigCategory = '" & strL & "'"
        If StrComp(strL, "Global Settings", CompareMethod.Text) = 0 Then
            strS = "CHARCONFIGTITLE ASC"
        Else
            strS = "INTORDER ASC"
        End If
        strS = "CHARCONFIGTITLE ASC"

        Dim dv1 As system.data.dataview = New DataView(tbl1, strF, strS, DataViewRowState.CurrentRows)
        dv1.AllowNew = False
        dv1.AllowDelete = False
        dgv.DataSource = dv1

        dgv.Columns("charConfigTitle").ReadOnly = True
        dgv.Columns("Example").ReadOnly = True
        If StrComp(str1, "Directory Paths", CompareMethod.Text) = 0 Then
            dgv.Columns("charConfigValue").ReadOnly = True
        ElseIf StrComp(str1, "Password Settings", CompareMethod.Text) = 0 Then
            dgv.Columns("charConfigValue").ReadOnly = False
        End If

        'select first entry in dgv
        If dv1.Count = 0 Then
            GoTo end1
        End If

        dgv.Rows(0).Cells("charConfigTitle").Selected = True
        'determine if cmd browse needs to be displayed
        boolDir = dv1(0).Item("BOOLISDIRECTORY")
        boolFile = dv1(0).Item("BOOLISFILE")
        If boolDir Or boolFile Then
            Call SetGlobalBrowse()
            Me.cmdBrowseGlobal.Visible = True
        Else
            Me.cmdBrowseGlobal.Visible = False
        End If

        dgv.Columns("Example").Visible = False

        'set dgv bool comboboxes
        For Count1 = 0 To dgv.RowCount - 1
            int2 = dgv("BOOLISBOOLEAN", Count1).Value
            If int2 = -1 Then
                var1 = dgv("CHARCONFIGTITLE", Count1).Value 'debug
                Dim cbx3 As New DataGridViewComboBoxCell
                cbx3.Items.Add("TRUE")
                cbx3.Items.Add("FALSE")
                'str1 = "Allow users to exclude data in StudyDoc (Watson overrides StudyDoc)"
                'int1 = FindRowDVByCol(str1, dv1, "CHARCONFIGTITLE")
                dgv("CHARCONFIGVALUE", Count1) = cbx3
                'show Example column
                dgv.Columns("Example").Visible = False
                var1 = dgv("CHARCONFIGVALUE", Count1).Value
                cbx3.Value = CStr(var1)
            End If

        Next

        'if strL = global then make Date column a dropdownbox
        If StrComp(strL, "Global Settings", CompareMethod.Text) = 0 Then
            Dim cbx As New DataGridViewComboBoxCell
            cbx.DataSource = tblDateFormats.Select("ID_TBLDATEFORMATS > 0", "INTORDER ASC")
            cbx.DisplayMember = tblDateFormats.Columns("CHARFORMAT").ColumnName
            'int1 = FindRowDV("Date Format", dv1)
            int1 = FindRowDVByCol("Table Date Format", dv1, "CHARCONFIGTITLE")
            dgv("CHARCONFIGVALUE", int1) = cbx
            'show Example column
            dgv.Columns("Example").Visible = True
            'fill with date example
            Dim tbl2 As System.Data.DataTable
            Dim rows2() As DataRow

            tbl2 = tblDateFormats
            var1 = dgv("CHARCONFIGVALUE", int1).Value
            strF = "CHARFORMAT = '" & var1 & "'"
            rows2 = tbl2.Select(strF, "INTORDER ASC")
            If rows2.Length = 0 Then
            Else
                str1 = NZ(rows2(0).Item("CHARDESCRIPTION"), "")
                dgv("Example", int1).Value = str1 & " for Sep 1, 2006"
            End If

            Dim cbx1 As New DataGridViewComboBoxCell
            cbx1.DataSource = tblDateFormats.Select("ID_TBLDATEFORMATS > 0", "INTORDER ASC")
            cbx1.DisplayMember = tblDateFormats.Columns("CHARFORMAT").ColumnName
            'int1 = FindRowDV("Date Format", dv1)
            int1 = FindRowDVByCol("Text Date Format", dv1, "CHARCONFIGTITLE")
            dgv("CHARCONFIGVALUE", int1) = cbx1
            'show Example column
            dgv.Columns("Example").Visible = True
            'fill with date example
            var1 = dgv("CHARCONFIGVALUE", int1).Value
            strF = "CHARFORMAT = '" & var1 & "'"
            rows2 = tbl2.Select(strF, "INTORDER ASC")
            If rows2.Length = 0 Then
            Else
                str1 = NZ(rows2(0).Item("CHARDESCRIPTION"), "")
                dgv("Example", int1).Value = str1 & " for Sep 1, 2006"
            End If

            Dim cbx2 As New DataGridViewComboBoxCell
            cbx2 = cbxxIncSmplDiff.Clone
            int1 = FindRowDVByCol("Default Incurred Sample %Diff Calculation", dv1, "CHARCONFIGTITLE")
            dgv("CHARCONFIGVALUE", int1) = cbx2
            'show Example column
            dgv.Columns("Example").Visible = True
            'fill with date example
            var1 = dgv("CHARCONFIGVALUE", int1).Value
            If StrComp(var1, "%Difference", CompareMethod.Text) = 0 Then
                str1 = "(Incurred - Original)/Original * 100"
            ElseIf StrComp(var1, "Mean %Difference", CompareMethod.Text) = 0 Then
                str1 = "(Incurred - Original)/Mean * 100"
            End If
            dgv("Example", int1).Value = str1


        Else
            'dgv.Columns("Example").Visible = False
        End If

        'check for alignment
        If StrComp(strL, "Directory Paths", CompareMethod.Text) = 0 Then
            dgv.Columns("charConfigValue").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
        Else
            dgv.Columns("charConfigValue").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        End If

        dgv.AutoResizeColumns()


end1:
    End Sub

    Sub FillUserboolA(ByVal dgv As DataGridView)

        Dim Count1 As Short
        Dim intRows As Short
        Dim int1 As Short

        boolHold = True

        intRows = dgv.RowCount

        Dim dv As system.data.dataview
        dv = dgv.DataSource

        Try
            For Count1 = 0 To intRows - 1
                int1 = dgv("BOOLACTIVE", Count1).Value
                If int1 = 0 Then
                    dgv("BOOLA", Count1).Value = False
                Else
                    dgv("BOOLA", Count1).Value = True
                End If
                'dv.Item(Count1).BeginEdit()
                'If int1 = -1 Then
                '    'dgv("BOOLA", Count1).Value = True
                '    dv.Item(Count1).Item("BOOLA") = True
                'Else
                '    'dgv("BOOLA", Count1).Value = False
                '    dv.Item(Count1).Item("BOOLA") = False
                'End If
                'dv.Item(Count1).EndEdit()
            Next
        Catch ex As Exception

        End Try

        boolHold = False

    End Sub

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        Call DoThisAdmin("Edit")
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        boolSave = True
        Call DoThisAdmin("Save")
        boolSave = False

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Call DoThisAdmin("Cancel")
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click

        Call DoExit()

    End Sub

    Sub DoExit()

        Me.Visible = False

        If StrComp(frmName, "frmHome_01", CompareMethod.Text) = 0 Then
            frmH.Visible = True
        ElseIf StrComp(frmName, "frmConsole", CompareMethod.Text) = 0 Then
            frmC.Visible = True
        End If

    End Sub

    Private Sub lbxTab1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbxTab1.SelectedIndexChanged

        If boolFormLoad Then
            Exit Sub
        End If
        Call lbxTab1Change()

    End Sub

    Private Sub dgvUsers_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvUsers.CellBeginEdit

        If boolFormLoad Then
            Exit Sub
        End If

        Try
            'do not allow change to aaAdmin
            Dim var1
            Dim strM As String

            var1 = NZ(Me.dgvUsers.Rows(e.RowIndex).Cells("charLastName").Value, "")
            If StrComp(var1, "aaAdmin", CompareMethod.Text) = 0 Then
                e.Cancel = True
                'MsgBox("The Admin user cannot be modified." & ChrW(10) & "(CellBeginEdit_Users)", MsgBoxStyle.Information, "Invalid entry...")
                strM = "The Admin user cannot be modified."
                MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
            End If
        Catch ex As Exception

        End Try



    End Sub

    Private Sub dgvUsers_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvUsers.CellContentClick

        'testing
        'Exit Sub

        Dim intRow As Short
        Dim intCol As Short
        Dim intColboolA As Short
        Dim str1 As String
        Dim str3 As String
        Dim str4 As String
        Dim dgv As DataGridView
        Dim bool As Boolean
        Dim int1 As Short

        If Me.cmdEdit.Enabled Then
            Exit Sub
        End If

        If boolHold Then
            Exit Sub
        End If

        dgv = Me.dgvUsers
        intRow = e.RowIndex
        intCol = e.ColumnIndex

        If dgv.RowCount <= intRow Then 'NDL: Row has been deleted, don't do anything.
            Exit Sub
        End If

        'don't allow change to aaAdmin
        Try
            str1 = dgv("CHARLASTNAME", intRow).Value
            int1 = dgv("boolActive", intRow).Value

            If StrComp(str1, "aaAdmin", CompareMethod.Text) = 0 Then
                If int1 = 0 Then
                    Dim str2 As String
                    str2 = "Admin account cannot be deactivated"
                    MsgBox(str2, MsgBoxStyle.Information, "Invalid action...")
                    'bool = dgv.Item(intCol, intRow).Value
                    'If bool Then
                    '    dgv.Item("boolActive", intRow).Value = -1
                    'Else
                    '    dgv.Item("boolActive", intRow).Value = 0
                    'End If
                    dgv.Item("boolActive", intRow).Value = -1
                    Exit Sub
                End If
            End If
        Catch ex As Exception

        End Try


    End Sub

    Private Sub dgvUsers_CellContextMenuStripChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvUsers.CellContextMenuStripChanged

    End Sub

    Private Sub dgvUsers_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvUsers.CellValueChanged

        'testing
        'Dim v
        'v = e.RowIndex
        'Exit Sub

        If boolFormLoad Then
            Exit Sub
        End If

        If boolHold Then
            Exit Sub
        End If

        Dim str1 As String
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim bool As Short
        Dim dt As Date
        Dim var1
        Dim strF As String
        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim intRow As Short
        Dim intCol As Short


        dgv = Me.dgvUsers

        intRow = dgv.CurrentRow.Index ' e.RowIndex
        intCol = dgv.CurrentCell.ColumnIndex 'e.ColumnIndex

        dv = dgv.DataSource
        str1 = dgv.Columns(e.ColumnIndex).Name
        If StrComp(str1, "boolActive", CompareMethod.Text) = 0 Then
            bool = dgv.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
            ''for some reason, modifying dv is generating an error when saving tblPersonnel
            'dv(e.RowIndex).BeginEdit()
            'If bool = -1 Then
            '    dv(e.RowIndex).Item("dtDeactivated") = System.DBNull.Value
            'Else
            '    dt = Now
            '    dv(e.RowIndex).Item("dtDeactivated") = dt
            'End If
            'dv(e.RowIndex).EndEdit()
            'dgv.EndEdit(True)

            var1 = dgv("ID_TBLPERSONNEL", e.RowIndex).Value
            strF = "ID_TBLPERSONNEL = " & var1
            dtbl = tblPersonnel
            rows = dtbl.Select(strF)
            rows(0).BeginEdit()
            'If bool = -1 Then
            '    rows(0).Item("dtDeactivated") = System.DBNull.Value
            'Else
            '    dt = Now
            '    rows(0).Item("dtDeactivated") = dt
            'End If

            'bool is returning 1, same as -1
            If bool = 0 Then
                dt = Now
                rows(0).Item("dtDeactivated") = dt
            Else
                rows(0).Item("dtDeactivated") = System.DBNull.Value
            End If

            rows(0).EndEdit()

        ElseIf StrComp(str1, "boolA", CompareMethod.Text) = 0 Then

            bool = dgv.Item(intCol, intRow).Value
            If bool Then
                dgv.Item("boolActive", intRow).Value = -1
            Else
                dgv.Item("boolActive", intRow).Value = 0
            End If



        End If

        'Call FillUserboolA(dgv)


    End Sub

    Private Sub rbShowAllUsers_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbShowAllUsers.CheckedChanged
        If boolFormLoad Then
            Exit Sub
        End If
        Dim int1 As Short
        int1 = Me.dgvUsers.CurrentRow.Index

        Call ShowUsers(int1)

    End Sub

    Private Sub rbShowActiveUsers_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbShowActiveUsers.CheckedChanged
        If boolFormLoad Then
            Exit Sub
        End If
        Dim int1 As Short
        int1 = Me.dgvUsers.CurrentRow.Index

        Call ShowUsers(int1)
    End Sub

    Private Sub rbShowInactiveUsers_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If boolFormLoad Then
            Exit Sub
        End If
        Dim int1 As Short

        'int1 = Me.dgvUsers.CurrentRow.Index

        If Me.dgvUsers.RowCount = 0 Then
            int1 = -1
        ElseIf Me.dgvUsers.CurrentRow Is Nothing Then
            int1 = 0
        Else
            int1 = Me.dgvUsers.CurrentRow.Index
        End If
        Call ShowUsers(int1)

    End Sub

    Private Sub rbShowAllUserIDs_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbShowAllUserIDs.CheckedChanged
        If boolFormLoad Then
            Exit Sub
        End If
        Dim int1 As Short
        If Me.dgvUserAttributes.Rows.Count = 0 Then
            int1 = 0
        Else
            int1 = Me.dgvUserAttributes.CurrentRow.Index
        End If
        Call ShowAccounts(int1)

    End Sub

    Private Sub rbShowActiveUserIDs_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbShowActiveUserIDs.CheckedChanged
        If boolFormLoad Then
            Exit Sub
        End If
        Dim int1 As Short
        If Me.dgvUserAttributes.Rows.Count = 0 Then
            int1 = 0
        Else
            int1 = Me.dgvUserAttributes.CurrentRow.Index
        End If
        Call ShowAccounts(int1)

    End Sub

    Private Sub rbShowInactiveUsers_CheckedChanged1(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbShowInactiveUsers.CheckedChanged
        If boolFormLoad Then
            Exit Sub
        End If
        Dim int1 As Short

        'int1 = Me.dgvUsers.CurrentRow.Index

        If Me.dgvUsers.RowCount = 0 Then
            int1 = -1
        ElseIf Me.dgvUsers.CurrentRow Is Nothing Then
            int1 = 0
        Else
            int1 = Me.dgvUsers.CurrentRow.Index
        End If
        Call ShowUsers(int1)

    End Sub

    Private Sub rbShowInactiveUserIDs_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbShowInactiveUserIDs.CheckedChanged
        If boolFormLoad Then
            Exit Sub
        End If
        Dim int1 As Short
        If Me.dgvUserAttributes.Rows.Count = 0 Then
            int1 = 0
        Else
            int1 = Me.dgvUserAttributes.CurrentRow.Index
        End If
        Call ShowAccounts(int1)

    End Sub

    Private Sub cmdResetUserAccounts_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdResetUserAccounts.Click
        Call DoCancelUserAccountTab()
    End Sub

    Private Sub cmdResetDropdownbox_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdResetDropdownbox.Click
        Call DoCancelDropdownboxTab()
    End Sub

    Private Sub cmdResetCorporateAddressses_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdResetCorporateAddressses.Click
        Call DoCancelCorporateAddressTab()
    End Sub

    Private Sub cmdResetDefineReports_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdResetDefineReports.Click
        Call DoCancelReportTemplatesTab()
    End Sub

    Private Sub cmdAddUser_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAddUser.Click

        Dim frm As New frmAddUser

        Dim Count1 As Int32
        Dim Count2 As Int32
        Dim dgv As DataGridView
        Dim id1 As Int64

        frm.strForm = "UserName"
        Call frm.FormLoad()
        frm.Text = "Enter New User..."

        'place the form
        Dim a, b, c, d

        a = Me.pan1.Top
        b = Me.dgvUsers.Top
        c = Me.pan1.Left + Me.cmdAddUser.Left

        frm.Left = c
        frm.Top = a + b

        frm.txtFirstName.Focus()

        frm.ShowDialog()

        If frm.boolCancel Then
            frm.Dispose()
            GoTo end1
        End If

        Dim strFirstName As String
        Dim strMiddleName As String
        Dim strLastName As String

        strFirstName = frm.txtFirstName.Text
        strMiddleName = frm.txtMiddleName.Text
        strLastName = frm.txtLastName.Text

        frm.Dispose()

        'unselect all user rows
        'to avoid a validation error
        dgv = Me.dgvUsers
        For Count1 = 0 To dgv.Rows.Count - 1
            For Count2 = 0 To dgv.Columns.Count - 1
                dgv.Rows(Count1).Cells(Count2).Selected = False
            Next
            dgv.Rows(Count1).Selected = False
        Next

        Dim dv As System.Data.DataView
        Dim dt As Date
        Dim dtbl As System.Data.DataTable

        Dim tblMax As System.Data.DataTable
        Dim rowsMax() As DataRow
        Dim strFMax As String
        Dim maxID

        maxID = GetMaxID("TBLPERSONNEL", 1, True)
        '20190219 LEE: Don't need anymore. Used GetMaxID
        'Call PutMaxID("TBLPERSONNEL", maxID)

        boolAddUser = True

        'first switch view to Show All
        Me.rbShowAllUsers.Checked = True

        dtbl = tblPersonnel
        dt = Now

        Dim nr As DataRow = dtbl.NewRow
        nr.BeginEdit()
        nr.Item("boolActive") = -1 'True
        'dt will get updated later
        'must put a placeholder date to account for cellvalidate and boolActive
        nr.Item("dtActivated") = dt ' Format(dt, "MM/dd/yyyy")
        nr.Item("id_tblPersonnel") = maxID
        nr.Item("CHARFIRSTNAME") = strFirstName
        nr.Item("CHARMIDDLENAME") = strMiddleName
        nr.Item("CHARLASTNAME") = strLastName
        nr.Item("boolA") = True

        nr.EndEdit()
        dtbl.Rows.Add(nr)


        dv = dtbl.DefaultView
        dv.AllowNew = False
        dv.AllowDelete = False
        Me.dgvUsers.CurrentCell = Me.dgvUsers.Rows(0).Cells("charLastName")
        Me.dgvUsers.AutoResizeColumns()

        'clear checkboxes 
        Call DVToPasswordCheckboxValues()

        Call FillUserboolA(Me.dgvUsers)

        'clear dgvUserAccounts
        dv = Me.dgvUserAttributes.DataSource
        dv.RowFilter = "id_tblPersonnel = 0"
        Me.dgvUserAttributes.AutoResizeColumns()


        'unselect all stuff again
        For Count1 = 0 To dgv.Rows.Count - 1
            For Count2 = 0 To dgv.Columns.Count - 1
                dgv.Rows(Count1).Cells(Count2).Selected = False
            Next
            dgv.Rows(Count1).Selected = False
        Next

        'select new row
        For Count1 = 0 To dgv.Rows.Count - 1
            id1 = dgv("ID_TBLPERSONNEL", Count1).Value
            If id1 = maxID Then
                'select this row
                dgv.Rows(Count1).Selected = True
                'select the first cell as well
                dgv.Rows(Count1).Cells("CHARLASTNAME").Selected = True
            End If
        Next

        Dim intRow As Int32
        intRow = dgv.CurrentRow.Index
        dgv("boolActive", intRow).Value = -1 ' True

        'Call dgvUserSelectionChange()

end1:

        boolAddUser = False

    End Sub

    Private Sub dgvUserAttributes_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvUserAttributes.CellBeginEdit
        Dim var1

        If boolFormLoad Then
            Exit Sub
        End If

        var1 = NZ(Me.dgvUserAttributes.Rows(e.RowIndex).Cells("charuserid").Value, "")
        If StrComp(var1, "aaAdmin", CompareMethod.Text) = 0 Then
            e.Cancel = True
            MsgBox("The Admin user cannot be modified." & ChrW(10) & "(CellBeginEdit_UserAttributes)", MsgBoxStyle.Information, "Invalid entry...")
        End If

    End Sub

    Private Sub dgvUserAttributes_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvUserAttributes.CellValueChanged
        If boolFormLoad Then
            Exit Sub
        End If

        If boolHold Then
            Exit Sub
        End If

        Dim str1 As String
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim bool As Short
        Dim dt As Date
        Dim var1
        Dim strF As String
        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim intRow As Short
        Dim intCol As Short


        dgv = Me.dgvUserAttributes

        intRow = dgv.CurrentRow.Index ' e.RowIndex
        intCol = dgv.CurrentCell.ColumnIndex 'e.ColumnIndex

        dv = dgv.DataSource
        str1 = dgv.Columns(e.ColumnIndex).Name
        If StrComp(str1, "boolActive", CompareMethod.Text) = 0 Then
            bool = dgv.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
            ''for some reason, modifying dv is generating an error when saving tblPersonnel
            'dv(e.RowIndex).BeginEdit()
            'If bool = -1 Then
            '    dv(e.RowIndex).Item("dtDeactivated") = System.DBNull.Value
            'Else
            '    dt = Now
            '    dv(e.RowIndex).Item("dtDeactivated") = dt
            'End If
            'dv(e.RowIndex).EndEdit()
            'dgv.EndEdit(True)

            var1 = dgv("ID_TBLUSERACCOUNTS", e.RowIndex).Value
            strF = "ID_TBLUSERACCOUNTS = " & var1
            dtbl = tblUserAccounts
            rows = dtbl.Select(strF)
            rows(0).BeginEdit()

            'bool may return 1, which is the same as 0
            If bool = 0 Then
                dt = Now
                rows(0).Item("dtDeactivated") = dt
            Else
                rows(0).Item("dtDeactivated") = System.DBNull.Value
            End If

            rows(0).EndEdit()

        ElseIf StrComp(str1, "boolA", CompareMethod.Text) = 0 Then

            bool = dgv.Item(intCol, intRow).Value
            If bool Then
                dgv.Item("boolActive", intRow).Value = -1
            Else
                dgv.Item("boolActive", intRow).Value = 0
            End If



        End If

        'Call FillUserboolA(dgv)

    End Sub

    Private Sub dgvUserAttributes_CurrentCellDirtyStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvUserAttributes.CurrentCellDirtyStateChanged

        Dim intRow As Short
        Dim intCol As Short
        Dim str1 As String
        Dim dgv As DataGridView
        Dim bool As Boolean

        If Me.cmdEdit.Enabled Then
            Exit Sub
        End If

        If boolHold Then
            Exit Sub
        End If

        dgv = Me.dgvUserAttributes
        intRow = dgv.CurrentRow.Index ' e.RowIndex
        intCol = dgv.CurrentCell.ColumnIndex 'e.ColumnIndex

        str1 = dgv.Columns(intCol).Name
        If StrComp(str1, "boolA", CompareMethod.Text) = 0 Then
            dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If

    End Sub

    Private Sub dgvUserAttributes_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvUserAttributes.MouseEnter
        Me.dgvUserAttributes.Focus()
    End Sub

    Private Sub dgvUserAttributes_RowValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvUserAttributes.RowValidating
        If boolHold Then
            Exit Sub
        End If

        'Note:
        'if boolA is changed of the last row of dgv
        'rowfilter changes contents of dgv
        'but rowfilter hasn't fired yet
        'get an error when checking for unique name combinations
        'resolve by evaluating id_tblPersonnel
        'id_tblPersonnel will be nothing during rowfilter transition
        Dim dgv As DataGridView = Me.dgvUserAttributes
        Dim intRow As Int32 = e.RowIndex
        Dim vID
        vID = dgv("id_tblUserAccounts", intRow).Value

        If IsDBNull(vID) Or IsNothing(vID) Then
            Exit Sub
        End If

        If UserAccountValidate() Then
            If PasswordValidate() Then
            Else
                e.Cancel = True
            End If
        Else
            e.Cancel = True
        End If


    End Sub



    Private Sub dgvUsers_CurrentCellDirtyStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvUsers.CurrentCellDirtyStateChanged

        'testing
        'Exit Sub


        Dim intRow As Short
        Dim intCol As Short
        Dim str1 As String
        Dim dgv As DataGridView
        Dim bool As Boolean

        If Me.cmdEdit.Enabled Then
            Exit Sub
        End If

        If boolHold Then
            Exit Sub
        End If

        dgv = Me.dgvUsers
        intRow = dgv.CurrentRow.Index ' e.RowIndex
        intCol = dgv.CurrentCell.ColumnIndex 'e.ColumnIndex

        str1 = dgv.Columns(intCol).Name
        If StrComp(str1, "boolA", CompareMethod.Text) = 0 Then
            dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If

    End Sub

    Private Sub dgvUsers_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles dgvUsers.DataError

        'MsgBox("StudyDoc:  " & e.Context.ToString)

    End Sub


    Private Sub dgvUsers_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvUsers.MouseEnter

        'testing
        'Exit Sub

        Me.dgvUsers.Focus()
    End Sub

    Private Sub dgvUsers_RowValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvUsers.RowValidating

        'Exit Sub

        'testing
        'Exit Sub

        Dim intRow As Short
        Dim intCol As Short
        Dim str1 As String
        Dim str2 As String
        Dim var1, var2, var3, var4, var5
        Dim col As DataGridViewColumn
        Dim dgv As DataGridView
        Dim dtbl As System.Data.DataTable
        Dim ct1 As Short
        Dim Count1 As Short
        Dim strMsg As String

        If boolFormLoad Or boolAddUser Then
            Exit Sub
        End If

        If Me.cmdEdit.Enabled Then
            Exit Sub
        End If

        dgv = Me.dgvUsers
        str1 = dgv.Columns(intCol).Name
        intRow = e.RowIndex

        'Note:
        'if boolA is changed of the last row of dgv
        'rowfilter changes contents of dgv
        'but rowfilter hasn't fired yet
        'get an error when checking for unique name combinations
        'resolve by evaluating id_tblPersonnel
        'id_tblPersonnel will be nothing during rowfilter transition
        Dim vID
        vID = dgv("id_tblPersonnel", intRow).Value

        If IsDBNull(vID) Or IsNothing(vID) Then
            Exit Sub
        End If

        For Each col In dgv.Columns
            If StrComp(col.Name, "charLastName", CompareMethod.Text) = 0 Then
                var1 = dgv(col.Index, intRow).Value
                If Len(Trim(NZ(var1, ""))) = 0 Then
                    e.Cancel = True
                    dgv.CurrentCell = dgv.Rows(intRow).Cells(col.Index)
                    If boolAddUser Then
                    Else
                        MsgBox("The Last Name field cannot be blank,", MsgBoxStyle.Information, "Invalid entry...")
                    End If
                    Exit For
                End If
            End If
            If StrComp(col.Name, "charFirstName", CompareMethod.Text) = 0 Then
                var1 = dgv(col.Index, intRow).Value
                If Len(Trim(NZ(var1, ""))) = 0 Then
                    e.Cancel = True
                    dgv.CurrentCell = dgv.Rows(intRow).Cells(col.Index)
                    If boolAddUser Then
                    Else
                        MsgBox("The First Name field cannot be blank,", MsgBoxStyle.Information, "Invalid entry...")
                    End If
                    Exit For
                End If
            End If
        Next

        'now ensure that First/Last Name combo are unique


        dtbl = tblPersonnel
        ct1 = dtbl.Rows.Count
        var1 = NZ(dgv("charLastName", e.RowIndex).Value, "")
        var2 = NZ(dgv("charFirstName", e.RowIndex).Value, "")
        var5 = NZ(dgv("charMiddleName", e.RowIndex).Value, "")
        var4 = dgv("id_tblPersonnel", e.RowIndex).Value
        str1 = var1 & var2 & var5
        For Count1 = 0 To ct1 - 1
            Try
                var3 = dtbl.Rows(Count1).Item("id_tblPersonnel")
                If var3 = var4 Then 'ignore
                Else
                    'var1 = NZ(dgv("charLastName", Count1).Value, "")
                    'var2 = NZ(dgv("charFirstName", Count1).Value, "")
                    'var5 = NZ(dgv("charMiddleName", Count1).Value, "")
                    var1 = NZ(dtbl.Rows(Count1).Item("charLastName"), "")
                    var2 = NZ(dtbl.Rows(Count1).Item("charFirstName"), "")
                    var5 = NZ(dtbl.Rows(Count1).Item("charMiddleName"), "")
                    str2 = var1 & var2 & var5
                    If StrComp(str1, str2, CompareMethod.Text) = 0 Then
                        e.Cancel = True
                        var3 = dtbl.Rows(Count1).Item("boolActive")
                        If var3 = 0 Then
                            strMsg = "An inactive account named " & var2 & " " & var5 & " " & var1 & " already exists."
                        Else
                            strMsg = "An active account named " & var2 & " " & var5 & " " & var1 & " already exists."
                        End If
                        MsgBox(strMsg, MsgBoxStyle.Information, "Invalid entry...")
                        Exit For
                    End If
                End If
            Catch ex As Exception

            End Try
        Next

    End Sub

    Private Sub cmdAddUserID_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAddUserID.Click

        If Me.dgvUsers.Rows.Count = 0 Then
            Exit Sub
        End If

        'get idP
        Dim int1 As Short

        int1 = Me.dgvUsers.CurrentRow.Index

        Dim Count1 As Int32
        Dim Count2 As Short
        Dim dgv As DataGridView
        Dim id1 As Int64
        Dim idP As Int64
        Dim str1 As String

        idP = Me.dgvUsers("ID_TBLPERSONNEL", int1).Value

        dgv = Me.dgvUserAttributes

        Dim frm As New frmAddUser

        frm.strForm = "UserID"
        Call frm.FormLoad()
        frm.Text = "Enter New UserID..."
        frm.idP = idP
        'place the form
        Dim a, b, c, d

        a = Me.pan1.Top
        b = Me.dgvUsers.Top
        c = Me.pan1.Left + Me.cmdAddUser.Left

        frm.Left = c
        frm.Top = a + b

        frm.txtUserID.Focus()

        frm.ShowDialog()

        If frm.boolCancel Then
            frm.Dispose()
            GoTo end1
        End If

        Dim strUserID As String
        Dim strPassword As String

        strUserID = frm.txtUserID.Text
        str1 = frm.txtPswd.Text
        strPassword = PasswordEncrypt(str1) ' Decode(Coding(str1, True), False)

        frm.Dispose()

        'unselect all user rows
        'to avoid a validation error
        For Count1 = 0 To dgv.Rows.Count - 1
            For Count2 = 0 To dgv.Columns.Count - 1
                dgv.Rows(Count1).Cells(Count2).Selected = False
            Next
            dgv.Rows(Count1).Selected = False
        Next

        Dim strM As String

        Dim dv As system.data.dataview

        Dim dt As Date
        Dim dtbl As System.Data.DataTable
        Dim var1
        Dim strS As String
        Dim strF As String

        Dim tblMax As System.Data.DataTable
        Dim rowsMax() As DataRow
        Dim strFMax As String
        Dim maxID, maxIDPerm

        maxID = 1
        maxID = GetMaxID("TBLUSERACCOUNTS", 1, True)
        '20190219 LEE: Don't need anymore. Used GetMaxID
        'Call PutMaxID("TBLUSERACCOUNTS", maxID)

        'record password information
        ctPswd = ctPswd + 1
        If ctPswd > UBound(arrPswd, 2) Then
            ReDim Preserve arrPswd(10, UBound(arrPswd, 2) + 100)
        End If
        arrPswd(1, ctPswd) = ctPswd
        arrPswd(2, ctPswd) = CLng(maxID) 'this is id_tblUserAccounts
        arrPswd(3, ctPswd) = strPassword

        boolAddAcct = True

        'increment place holder number
        'incr1 = incr1 + 1

        boolAddUserAccount = True

        'switch view to Show All
        Me.rbShowAllUserIDs.Checked = True

        strF = "id_tblPersonnel = " & idP
        strS = "charUserID ASC"
        dtbl = tblUserAccounts
        dt = Now

        Dim nr As DataRow = dtbl.NewRow
        nr.BeginEdit()
        nr.Item("id_tblUserAccounts") = maxID
        nr.Item("CHARUSERID") = strUserID
        nr.Item("CHARPASSWORD") = strPassword
        nr.Item("boolActive") = -1 'True
        nr.Item("DTTIMESTAMP") = dt
        nr.Item("UPSIZE_TS") = dt
        'dt will get updated later
        'must put a placeholder date to account for cellvalidate and boolActive
        nr.Item("dtActivated") = dt ' Format(dt, "MM/dd/yyyy")
        nr.Item("BOOLCHANGEPASSWORDATNEXTLOGON") = -1
        nr.Item("BOOLUSERCANNOTCHANGEPASSWORD") = 0
        nr.Item("BOOLPASSWORDNEVEREXPIRES") = 0
        nr.Item("BOOLACCOUNTISLOCKEDOUT") = 0
        nr.Item("CHARPASSWORD") = strPassword
        nr.Item("ID_TBLPERMISSIONS") = 1
        nr.Item("ID_TBLPERSONNEL") = idP
        nr.Item("ID_TBLWATSONACCOUNT") = 0
        nr.Item("ID_TBLWINDOWSAUTH") = 0
        'don't need to add anything to CHARLDAP
        nr.Item("boolA") = True

        'nr.Item("numIncr") = incr1
        'DO NOT ADD dtTimeStamp!
        'Will be added later!
        ' nr.Item("dtTimeStamp") = dt
        nr.EndEdit()
        dtbl.Rows.Add(nr)

        dv = New DataView(dtbl)

        dv.AllowNew = False
        dv.AllowDelete = False
        dv.RowFilter = strF
        dv.Sort = strS
        Me.dgvUserAttributes.Columns("charUserID").Frozen = False
        Me.dgvUserAttributes.DataSource = dv
        Me.dgvUserAttributes.CurrentCell = Me.dgvUserAttributes.Rows(0).Cells("charUserID")
        Me.dgvUserAttributes.AutoResizeColumns()

        dgv.DataSource = dv

        Try
            'Me.dgvUserAttributes.Columns("charUserID").Frozen = True
        Catch ex As Exception

        End Try

        boolAddUserAccount = False

        Call DVToPasswordCheckboxValues()

        Call EvaluateUserAccounts()

        For Count1 = 0 To dgv.Rows.Count - 1
            id1 = dgv("ID_TBLUSERACCOUNTS", Count1).Value
            If id1 = maxID Then
                'select this row
                dgv.Rows(Count1).Selected = True
            End If
        Next

        'check boolActive in current row
        'for some reason, it isn't checking

        'unselect all userid rows AGAIN!
        'to avoid a validation error
        For Count1 = 0 To dgv.Rows.Count - 1
            For Count2 = 0 To dgv.Columns.Count - 1
                dgv.Rows(Count1).Cells(Count2).Selected = False
            Next
            dgv.Rows(Count1).Selected = False
        Next

        'select new row
        For Count1 = 0 To dgv.Rows.Count - 1
            id1 = dgv("ID_TBLUSERACCOUNTS", Count1).Value
            If id1 = maxID Then
                'select this row
                dgv.Rows(Count1).Selected = True
                'select the first cell as well
                dgv.Rows(Count1).Cells("CHARUSERID").Selected = True

                Dim dgvr As DataGridViewRow = dgv.Rows(Count1)
                dgv.CurrentCell = dgv.Rows(Count1).Cells("CHARUSERID")

            End If
        Next


        'pesky
        Dim intRow As Int32
        intRow = Me.dgvUserAttributes.CurrentRow.Index
        Me.dgvUserAttributes("boolActive", intRow).Value = -1 ' True

        'make admin default permissions group
        str1 = "Administrator"
        Me.cbxPermissionsGroup.SelectedIndex = Me.cbxPermissionsGroup.FindString(str1)

        boolAddAcct = False

end1:


    End Sub

    Private Sub cmdEnterPassword_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdEnterPassword.Click

        Dim intRow As Short
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim var1

        dgv = Me.dgvUserAttributes
        If dgv.Rows.Count = 0 Then
            Exit Sub
        End If
        intRow = Me.dgvUserAttributes.CurrentRow.Index
        var1 = NZ(Me.dgvUserAttributes.Rows(intRow).Cells("charPassword").Value, "")
        If Len(var1) = 0 Then
        Else
            var1 = PasswordUnEncrypt(var1.ToString) ' Coding(Decode(var1, True), False)
        End If

        'Dim frm As New frmEnterPassword
        Dim frm As New frmPasswordChange

        frm.chkFromAdmin.Checked = True
        frm.chkFromChgPswd.Checked = False

        'MsgBox(frmPasswordChange.boolFromAdmin)'debug


        frm.txtOldPassword.Text = var1
        frm.txtOldPassword.Enabled = False

        var1 = Me.dgvUserAttributes.Rows(intRow).Cells("id_tblUserAccounts").Value
        frm.txtID.Text = ""
        frm.txtID.Text = var1

        frm.ShowDialog()
        'Me.Refresh()
        If frm.boolCancel Then
        Else
            'enter password in field
            dv = dgv.DataSource
            dv(intRow).BeginEdit()
            'var1 = frm.txtConfirmPassword.Text
            'var1 = Coding(var1, True)
            'var1 = Decode(var1, False)
            dv(intRow).Item("charPassword") = frm.strPswd 'var1
            dv(intRow).Item("DTTIMESTAMP") = Now
            dv(intRow).EndEdit()

            'get information from frm
            ctPswd = ctPswd + 1
            If ctPswd > UBound(arrPswd, 2) Then
                ReDim Preserve arrPswd(10, UBound(arrPswd, 2) + 100)
            End If
            arrPswd(1, ctPswd) = ctPswd
            arrPswd(2, ctPswd) = CLng(frm.txtID.Text)
            arrPswd(3, ctPswd) = frm.strPswd
        End If

        frm.Dispose()

    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim idP As Int64
        Dim idUA As Int64
        Dim idPerm As Int64
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim strP As String
        Dim strPerm As String
        Dim strUID As String
        Dim strF As String
        Dim strM As String

        idP = id_tblPersonnel
        idUA = id_tblUserAccounts
        idPerm = id_tblPermissions

        strF = "id_tblPersonnel = " & idP
        Dim rows() As DataRow = tblPersonnel.Select(strF)

        str1 = rows(0).Item("CHARFIRSTNAME")
        str2 = rows(0).Item("CHARLASTNAME")
        strP = str1 & " " & str2

        strF = "id_tblUserAccounts = " & idUA
        rows = tblUserAccounts.Select(strF)
        strUID = rows(0).Item("CHARUSERID")

        strF = "id_tblPermissions = " & idPerm
        rows = tblPermissions.Select(strF)
        strPerm = rows(0).Item("CHARPERMISSIONSNAME")

        strM = "idP = " & idP
        strM = strM & ChrW(10) & "idUA = " & idUA
        strM = strM & ChrW(10) & "idPerm = " & idPerm
        strM = strM & ChrW(10) & "UserName = " & strP
        strM = strM & ChrW(10) & "UserID = " & strUID
        strM = strM & ChrW(10) & "Permissions Name = " & strPerm

        MsgBox(strM)

    End Sub

    Private Sub dgvTemplateAttributes_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvTemplateAttributes.CellContentClick

        'fill boolActive column
        Dim str1 As String
        Dim dgv As DataGridView
        Dim var1
        Dim bool As Boolean
        Dim dv As system.data.dataview

        If e.RowIndex < 0 Then
            Exit Sub
        End If

        dgv = Me.dgvTemplateAttributes

        dv = dgv.DataSource
        str1 = dgv.Columns(e.ColumnIndex).Name
        If StrComp(str1, "boolI", CompareMethod.Text) = 0 Then
            dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
            var1 = dgv.Rows(e.RowIndex).Cells("boolI").Value
            If IsDBNull(var1) Then
            Else
                dv(e.RowIndex).BeginEdit()
                If var1 Then
                    dv(e.RowIndex).Item("boolInclude") = -1
                Else
                    dv(e.RowIndex).Item("boolInclude") = 0
                End If
                dv(e.RowIndex).EndEdit()

            End If
        End If

    End Sub

    Private Sub dgvTemplateAttributes_CellValueNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs) Handles dgvTemplateAttributes.CellValueNeeded
        Dim dgv As DataGridView
        Dim intRow As Short
        Dim intCol As Short
        Dim str1 As String
        Dim var1, var2
        Dim dv As system.data.dataview
        Dim row() As DataRow
        Dim strF As String
        Dim dtblT As System.Data.DataTable

        boolAddRow = True

        dgv = Me.dgvTemplateAttributes
        intCol = e.ColumnIndex
        intRow = e.RowIndex
        str1 = dgv.Columns(intCol).Name
        dv = dgv.DataSource
        'fill TabName
        If StrComp(str1, "TabName", CompareMethod.Text) = 0 Then
            var2 = e.Value
            dtblT = tblTab1
            Erase row
            var1 = dv(intRow).Item("id_tblTab1")
            strF = "id_tblTab1 = " & var1
            row = dtblT.Select(strF)
            var1 = row(0).Item("charItem")
            e.Value = var1
        End If

        boolAddRow = False

    End Sub

    Private Sub rbShowAllTemplates_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbShowAllTemplates.CheckedChanged
        Call ShowTemplate()
    End Sub

    Private Sub rbShowInactiveTemplates_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbShowInactiveTemplates.CheckedChanged
        Call ShowTemplate()
    End Sub

    Private Sub rbShowActiveTemplates_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbShowActiveTemplates.CheckedChanged
        Call ShowTemplate()
    End Sub

    Private Sub cmdAddTemplate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAddTemplate.Click

        Dim dtbl As System.Data.DataTable
        Dim dv As system.data.dataview
        Dim dv1 As system.data.dataview
        Dim int1 As Int32
        Dim strF As String
        Dim Count1 As Short
        Dim varID

        boolAddRow = True

        'first make sure all are shown
        Me.rbShowAllTemplates.Checked = True

        Dim tblMax As System.Data.DataTable
        Dim rowsMax() As DataRow
        Dim strFMax As String
        Dim maxID

        maxID = GetMaxID("TBLTEMPLATES", 1, True)
        '20190219 LEE: Don't need anymore. Used GetMaxID
        'Call PutMaxID("TBLTEMPLATES", maxID)

        incr1 = incr1 + 1

        dtbl = tblTemplates
        Dim nr As DataRow = dtbl.NewRow

        Dim boolF As Boolean = boolFormLoad

        boolFormLoad = True
        nr.BeginEdit()
        nr.Item("id_tblTemplates") = maxID
        nr.Item("boolActive") = -1 'True
        nr.Item("boolA") = True
        'nr.Item("numincr") = incr1
        nr.EndEdit()
        dtbl.Rows.Add(nr)
        int1 = dtbl.Rows.Count
        boolFormLoad = boolF

        'record primary key value
        varID = maxID 'dtbl.Rows(int1 - 1).Item("id_tblTemplates")

        dv = dtbl.DefaultView
        dv.AllowNew = False
        dv.AllowDelete = False
        Me.dgvTemplates.CurrentCell = Me.dgvTemplates.Rows(0).Cells("charTemplateName")
        Me.dgvTemplates.AutoResizeColumns()

        'add rows to dgvTemplateAttributes
        Dim tbl1 As System.Data.DataTable
        Dim rows() As DataRow
        dtbl = tblTemplateAttributes
        tbl1 = tblTab1
        strF = "boolincludeintemplate = -1 and (intForm = 1 or intForm = 2)" '2 for FieldCodes
        rows = tbl1.Select(strF)
        int1 = rows.Length
        For Count1 = 0 To int1 - 1
            Dim nr1 As DataRow = dtbl.NewRow
            nr1.BeginEdit()
            'nr1.Item("numincr") = incr1
            nr1.Item("id_tblTemplates") = varID
            nr1.Item("boolInclude") = -1
            nr1.Item("boolI") = True
            nr1.Item("id_tblTab1") = rows(Count1).Item("id_tblTab1")
            nr1.EndEdit()
            dtbl.Rows.Add(nr1)
        Next
        dv1 = dtbl.DefaultView
        strF = "id_tblTemplates = " & varID
        dv1.RowFilter = strF
        dv1.AllowDelete = False
        dv1.AllowNew = False
        Me.dgvTemplateAttributes.DataSource = dv1

        Call UpdateTabNames()

        'select dgvTemplates and go into edit mode

        Call UpdateStudyName()

        Me.dgvTemplates.Select()
        Me.dgvTemplates.CurrentCell = Me.dgvTemplates.Rows(0).Cells("charTemplateName")
        Me.dgvTemplates.BeginEdit(True)

        boolAddRow = False

    End Sub

    Private Sub dgvTemplates_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvTemplates.CellContentClick
        'fill boolActive column
        Dim str1 As String
        Dim dgv As DataGridView
        Dim var1
        Dim bool As Boolean
        Dim dv As system.data.dataview

        If e.RowIndex < 0 Then
            Exit Sub
        End If

        dgv = Me.dgvTemplates

        dv = dgv.DataSource
        str1 = dgv.Columns(e.ColumnIndex).Name
        If StrComp(str1, "boolA", CompareMethod.Text) = 0 Then
            dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)

            var1 = dgv.Rows(e.RowIndex).Cells("boolA").Value
            If IsDBNull(var1) Then
            Else
                dv(e.RowIndex).BeginEdit()
                If var1 Then
                    dv(e.RowIndex).Item("boolActive") = -1
                Else
                    dv(e.RowIndex).Item("boolActive") = 0
                End If
                dv(e.RowIndex).EndEdit()
            End If
        End If


    End Sub

    Private Sub dgvTemplates_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgvTemplates.CellValidating
        If boolFormLoad Then
            Exit Sub
        End If
        If Me.cmdEdit.Enabled Then
            Exit Sub
        End If
        If boolAddRow Then
            Exit Sub
        End If

        Dim intRow As Short
        Dim intCol As Short
        Dim dgv As DataGridView
        Dim var1, var2
        Dim str2 As String
        Dim boolTemplate As Boolean = False
        Dim boolStudyName As Boolean = False

        intRow = e.RowIndex
        intCol = e.ColumnIndex

        dgv = Me.dgvTemplates

        var1 = dgv.Columns(intCol).Name
        If StrComp(var1, "charTemplateName", CompareMethod.Text) = 0 Then 'continue
            boolTemplate = True
        ElseIf StrComp(var1, "StudyName", CompareMethod.Text) = 0 Then
            boolStudyName = True
        Else
            Exit Sub
        End If

        'some cells cannot be blank
        Dim dv As system.data.dataview
        Dim str1 As String
        Dim boolStop As Boolean
        Dim Count1 As Short
        Dim strT As String
        Dim int1 As Short

        dv = dgv.DataSource

        int1 = dgv.Columns.Count
        boolStop = False
        str2 = ""

        str1 = ""
        If boolTemplate Then
            var1 = e.FormattedValue ' dgv("charTemplateName", intRow).Value
            If Len(NZ(var1, "")) = 0 Then
                boolStop = True
                str1 = "charTemplateName"
                str2 = "Template Name"
            End If
        ElseIf boolStudyName Then
            var1 = e.FormattedValue 'dgv("StudyName", intRow).Value
            If Len(NZ(var1, "")) = 0 Then
                boolStop = True
                str1 = "StudyName"
                str2 = "Study Name"
            End If
        End If

        strT = str1
        If boolStop Then
            e.Cancel = True
            MsgBox(str2 & " cannot be blank.", MsgBoxStyle.Information, "Invalid entry...")
            dgv.Rows(intRow).Cells(strT).Selected = True
            dgv.BeginEdit(True)
            GoTo end1
        End If

        'now check for unique Template names
        If boolTemplate Then
            Dim bool As Short
            Dim strM As String
            Dim tbl As System.Data.DataTable

            tbl = tblTemplates
            ' str1 = dgv("charTemplateName", e.RowIndex).Value.ToString
            str1 = e.FormattedValue 'dv(e.RowIndex).Item("charTemplateName")
            var1 = dv(e.RowIndex).Item("id_tblTemplates")
            For Count1 = 0 To tbl.Rows.Count - 1
                var2 = tbl.Rows(Count1).Item("id_tblTemplates")
                If var1 = var2 Then 'ignore
                Else
                    str2 = tbl.Rows(Count1).Item("charTemplateName")
                    If StrComp(str1, str2, CompareMethod.Text) = 0 Then
                        bool = dgv("boolActive", Count1).Value
                        If bool = 0 Then
                            strM = "Inactive Template Name " & str2 & " already exists."
                        Else
                            strM = "Active Template Name " & str2 & " already exists."
                        End If
                        e.Cancel = True
                        MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
                        dgv.CurrentCell = dgv.Rows(e.RowIndex).Cells("charTemplateName")
                        Me.dgvTemplates.BeginEdit(True)
                        GoTo end1
                    End If
                End If
            Next
        End If

end1:


    End Sub

    Private Sub dgvTemplates_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvTemplates.CellValueChanged
        If boolFormLoad Then
            Exit Sub
        End If
        If Me.cmdEdit.Enabled Then
            'Exit Sub
        End If

        Dim int1 As Short
        Dim int2 As Short
        Dim dgv As DataGridView
        Dim str1 As String
        Dim str2 As String
        Dim strErr As String
        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim var1, var2

        dgv = Me.dgvTemplates
        int1 = e.ColumnIndex 'dgv.CurrentCell.ColumnIndex
        int2 = e.RowIndex 'dgv.CurrentRow.Index
        str2 = dgv.Columns(int1).Name

        If StrComp(str2, "StudyName", CompareMethod.Text) = 0 Then
            tbl = tblStudies
            'dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
            'enter contents of combobox into grid cell
            str1 = dgv("StudyName", int2).Value
            'find str1 in tbl
            strF = "charWatsonStudyName = '" & str1 & "'"
            rows = tbl.Select(strF)
            var1 = rows(0).Item("id_tblStudies")
            dgv("id_tblStudies", int2).Value = var1
            dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
            dgv.Update()
        End If

    End Sub

    Private Sub dgvTemplates_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvTemplates.MouseEnter
        Me.dgvTemplates.Focus()
    End Sub


    Private Sub dgvNickNames_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgvNickNames.CellValidating
        If boolAddAddresses Then
        Else
            Exit Sub
        End If
        If boolCancelAddresses Then
            Exit Sub
        End If
        If boolFormLoad Then
            Exit Sub
        End If
        If Me.cmdEdit.Enabled Then
            Exit Sub
        End If

        Dim dgv As DataGridView
        Dim dgv1 As DataGridView
        Dim var1, var2, var3
        Dim dv As system.data.dataview
        Dim int1 As Short
        Dim Count1 As Short
        Dim strF As String
        Dim boolGo As Boolean
        Dim strM As String
        Dim bool As Boolean

        dgv = Me.dgvNickNames
        var1 = dgv.Columns("charNickName").Index

        'ensure entry is not blank
        If var1 = e.ColumnIndex Then
            var2 = e.FormattedValue
            If IsDBNull(var1) Or Len(NZ(var2, "")) = 0 Then
                e.Cancel = True
                MsgBox("NickName cannot be blank.", MsgBoxStyle.Information, "Invalid Entry...")
                dgv.CurrentCell = dgv.Rows(e.RowIndex).Cells(e.ColumnIndex)
                dgv.BeginEdit(True)
            Else

                boolGo = True
                'check to see that value is unique
                strF = "charNickName = '" & e.FormattedValue.ToString & "'"
                Dim dtbl As System.Data.DataTable
                dtbl = tblCorporateNickNames
                int1 = dtbl.Rows.Count
                For Count1 = 0 To int1 - 1
                    var2 = dtbl.Rows(Count1).Item("id_tblCorporateNickNames")
                    var3 = dgv("id_tblCorporateNickNames", e.RowIndex).Value
                    If var2 = var3 Then 'ignore
                    Else 'investigate further
                        If StrComp(NZ(var1, ""), NZ(e.FormattedValue, ""), CompareMethod.Text) = 0 Then
                            e.Cancel = True
                            var2 = dtbl.Rows(Count1).Item("boolInclude")
                            If var2 = -1 Then
                                strM = "An active NickName of " & e.FormattedValue.ToString & " already exists."
                            Else
                                strM = "An inactive NickName of " & e.FormattedValue.ToString & " already exists."
                            End If
                            MsgBox(strM, MsgBoxStyle.Information, "Invalid Entry...")
                            dgv.CurrentCell = dgv.Rows(e.RowIndex).Cells(e.ColumnIndex)
                            dgv.BeginEdit(True)
                            boolGo = False
                            Exit For
                        Else

                        End If
                    End If

                Next

            End If

        End If

    End Sub

    Private Sub dgvNickNames_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvNickNames.Click
        Call CorporateAddressFilter()
    End Sub

    Private Sub cmdAddCorporateAddress_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAddCorporateAddress.Click

        Dim dtbl1 As System.Data.DataTable
        Dim dtbl2 As System.Data.DataTable
        Dim dtblL As System.Data.DataTable
        'Dim dv1 as system.data.dataview
        'Dim dv2 as system.data.dataview
        Dim rows() As DataRow
        Dim int1 As Short
        Dim int2 As Short
        Dim Count1 As Short
        Dim var1, var2
        Dim strF As String
        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView
        Dim strS As String

        Dim tblMax As System.Data.DataTable
        Dim rowsMax() As DataRow
        Dim strFMax As String
        Dim maxID

        maxID = 1

        maxID = GetMaxID("TBLCORPORATENICKNAMES", 1, True)
        '20190219 LEE: Don't need anymore. Used GetMaxID
        'Call PutMaxID("TBLCORPORATENICKNAMES", maxID)

        dtbl1 = tblCorporateNickNames
        dtbl2 = tblCorporateAddresses
        dtblL = tblAddressLabels

        dgv1 = Me.dgvNickNames

        'show all items first

        Dim boolF As Boolean = boolFormLoad

        boolFormLoad = True
        Me.rbShowAllAddresses.Checked = True
        boolFormLoad = boolF

        'add item to dtbl1
        incr1 = incr1 + 1
        Dim nr1 As DataRow = dtbl1.NewRow
        nr1.BeginEdit()
        'nr1("numincr") = incr1
        nr1("boolInclude") = -1 'True
        nr1("boolI") = True
        nr1("id_tblCorporateNickNames") = maxID
        nr1.EndEdit()
        dtbl1.Rows.Add(nr1)
        strF = "charNickName <> '[None]' OR charNickName is null"
        strS = "charNickName ASC"
        Dim dv1 As system.data.dataview = New DataView(dtbl1, strF, strS, DataViewRowState.CurrentRows)
        dv1.AllowNew = False
        dv1.AllowDelete = False
        'int2 = dv1.Count
        dgv1.DataSource = dv1

        'add items to dtbl2
        int1 = dtblL.Rows.Count
        For Count1 = 0 To int1 - 1
            var1 = dtblL.Rows(Count1).Item("id_tblAddressLabels")
            var2 = dtblL.Rows(Count1).Item("charLabel")
            Dim nr As DataRow = dtbl2.NewRow
            nr.BeginEdit()
            nr("id_tblCorporateNickNames") = maxID
            nr("numincr") = incr1
            nr("id_tblAddressLabels") = var1
            nr("charAddressLabel") = var2
            nr("boolIncludeInTitle") = False
            nr.EndEdit()
            dtbl2.Rows.Add(nr)
        Next

        'filter dgv2
        dgv2 = Me.dgvCorporateAddresses
        var1 = incr1
        'strF = "numincr = " & var1
        strF = "id_tblCorporateNickNames = " & maxID
        strS = "id_tblAddressLabels ASC"

        Dim dv2 As system.data.dataview = New DataView(dtbl2, strF, strS, DataViewRowState.CurrentRows)
        dv2.AllowNew = False
        dv2.AllowDelete = False
        dgv2.DataSource = dv2
        dgv2.AutoResizeColumns()

        'enter first cell in dgv1
        dgv1.BeginEdit(True)

        boolAddAddresses = True 'set this bool at the very end of the code


    End Sub

    Private Sub rbShowAllAddresses_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbShowAllAddresses.CheckedChanged
        Call ShowAddresses()
    End Sub

    Private Sub rbShowActiveAddresses_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbShowActiveAddresses.CheckedChanged
        Call ShowAddresses()
    End Sub

    Private Sub rbShowInactiveAddresses_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbShowInactiveAddresses.CheckedChanged
        Call ShowAddresses()
    End Sub


    Private Sub chkChangePasswordAtNextLogon_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkChangePasswordAtNextLogon.CheckedChanged

        If boolStopPswdCheck Then
            Exit Sub
        End If

        Call PasswordCheckboxes(False)
        Call PasswordCheckboxValuesToDV()

    End Sub

    Private Sub chkUserCannotChangePassword_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkUserCannotChangePassword.CheckedChanged

        If boolStopPswdCheck Then
            Exit Sub
        End If

        Call PasswordCheckboxValuesToDV()

    End Sub

    Private Sub chkPasswordNeverExpires_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkPasswordNeverExpires.CheckedChanged

        If boolStopPswdCheck Then
            Exit Sub
        End If

        Call PasswordCheckboxes(True)
        Call PasswordCheckboxValuesToDV()

    End Sub

    Private Sub chkAccountIsLockedOut_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAccountIsLockedOut.CheckedChanged

        If boolStopPswdCheck Then
            Exit Sub
        End If

        Call PasswordCheckboxValuesToDV()

    End Sub

    Private Sub cmdBrowseGlobal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdBrowseGlobal.Click

        Dim str1 As String
        Dim str2 As String
        Dim int1 As Short
        Dim strF As String
        Dim strPath As String
        Dim dgv As DataGridView
        Dim dv As System.Data.DataView
        Dim boolIsFile As Short
        Dim strConfigTitle As String

        dgv = Me.dgvGlobal
        int1 = dgv.CurrentRow.Index
        dv = dgv.DataSource
        boolIsFile = dv(int1).Item("BOOLISFILE")
        strConfigTitle = dv(int1).Item("CHARCONFIGTITLE")

        'get default path
        strPath = NZ(dv(int1).Item("charConfigValue"), "C:\")

        Dim strFilter As String
        Dim strFileName As String

        If InStr(1, strConfigTitle, "ChromReporter", CompareMethod.Text) > 0 Then
            strFilter = "ChromReporter.exe file (ChromeReporter.exe)|ChromReporter.exe"
            strFileName = "ChromReporter.exe"
        Else
            strFilter = "All files (*.*)|*.*"
            strFileName = "*.*"
        End If



        If boolIsFile = -1 Then
            str2 = ReturnDirectoryBrowse(True, strPath, strFilter, strFileName, True)
        Else
            str2 = ReturnDirectoryBrowse(False, strPath, strFilter, strFileName, True)
        End If


        If Len(str2) = 0 Then
        Else
            dv(int1).BeginEdit()
            dv(int1).Item("charConfigValue") = str2
            dv(int1).EndEdit()
        End If

    End Sub

    Private Sub lbxGlobal_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbxGlobal.SelectedIndexChanged

        Call GlobalConfigure()

        Dim str1 As String
        str1 = Me.lbxGlobal.SelectedItem
        If StrComp(str1, "Password Settings", CompareMethod.Text) = 0 Then
            Me.gbxlabelGlobalParameters.Visible = True
        Else
            Me.gbxlabelGlobalParameters.Visible = False
        End If

        Call SetGlobalParametersControls()

    End Sub

    Sub SetGlobalParametersControls()

        Dim a, b, c, d

        Dim h

        Dim bw As Int16 = (Me.Width - Me.ClientSize.Width) / 2 'form border width
        Dim tbh As Int16 = Me.Height - Me.ClientSize.Height - 2 * bw 'titlebar height

        h = Me.Height

        If Me.gbxlabelGlobalParameters.Visible Then
            a = Me.gbxlabelGlobalParameters.Top + Me.gbxlabelGlobalParameters.Height + 5

        Else
            a = Me.cmdResetGlobal.Top + Me.cmdResetGlobal.Height

        End If

        b = h - a - bw - 10

        Me.panGP.Top = a
        Me.panGP.Height = b


    End Sub

    Private Sub dgvGlobal_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgvGlobal.CellValidating

        If boolFormLoad Then
            Exit Sub
        End If
        If Me.cmdEdit.Enabled Then
            Exit Sub
        End If

        Dim str1 As String
        str1 = Me.dgvGlobal.Columns(e.ColumnIndex).Name
        If StrComp(str1, "CHARCONFIGVALUE", CompareMethod.Text) = 0 Then 'continue
        Else
            Exit Sub
        End If

        Dim dgv As DataGridView
        Dim boolN As Short
        Dim boolB As Short
        Dim bool As Boolean
        Dim var1
        str1 = Me.dgvGlobal.Rows(e.RowIndex).Cells("charConfigTitle").Value.ToString
        boolN = Me.dgvGlobal.Rows(e.RowIndex).Cells("boolIsNumeric").Value
        boolB = Me.dgvGlobal.Rows(e.RowIndex).Cells("boolIsBoolean").Value

        dgv = Me.dgvGlobal

        If boolN = -1 Then
            If StrComp(str1, "Default # of Decimals for Conc Data", CompareMethod.Text) = 0 Then 'entry must be integer >=0
                If IsNumeric(e.FormattedValue) Then 'continue
                    'number must be decimal
                    If IsInt(e.FormattedValue) Then
                        'number must be >= than 0
                        If e.FormattedValue < 0 Then
                            e.Cancel = True
                            MsgBox("Entry must be integer >= 0", MsgBoxStyle.Information, "Invalid entry...")
                        End If
                        'look for a decimal
                        If InStr(1, e.FormattedValue.ToString, ".", CompareMethod.Text) > 0 Then
                            e.Cancel = True
                            MsgBox("Entry must be integer >= 0", MsgBoxStyle.Information, "Invalid entry...")
                        End If
                    Else
                        e.Cancel = True
                        MsgBox("Entry must be integer >= 0", MsgBoxStyle.Information, "Invalid entry...")
                    End If
                Else
                    e.Cancel = True
                    MsgBox("Entry must be integer >= 0", MsgBoxStyle.Information, "Invalid entry...")
                End If
            Else 'entry must be integer > 0
                If IsNumeric(e.FormattedValue) Then 'continue
                    'number must be decimal
                    If IsInt(e.FormattedValue) Then
                        'number must be > than 0
                        If e.FormattedValue < 1 Then
                            e.Cancel = True
                            MsgBox("Entry must be integer > 0", MsgBoxStyle.Information, "Invalid entry...")
                        End If
                        'look for a decimal
                        If InStr(1, e.FormattedValue.ToString, ".", CompareMethod.Text) > 0 Then
                            e.Cancel = True
                            MsgBox("Entry must be integer > 0", MsgBoxStyle.Information, "Invalid entry...")
                        End If
                    Else
                        e.Cancel = True
                        MsgBox("Entry must be integer > 0", MsgBoxStyle.Information, "Invalid entry...")
                    End If
                Else
                    e.Cancel = True
                    MsgBox("Entry must be integer > 0", MsgBoxStyle.Information, "Invalid entry...")
                End If
            End If

        Else
            If boolB = -1 Then 'entry must be boolean
                If StrComp(NZ(e.FormattedValue, ""), "True", CompareMethod.Text) = 0 Or StrComp(NZ(e.FormattedValue, ""), "False", CompareMethod.Text) = 0 Then
                    'make all caps
                    str1 = e.FormattedValue
                    str1 = AllCaps(str1)
                    Dim dv As system.data.dataview
                    dv = Me.dgvGlobal.DataSource
                    dv(e.RowIndex).BeginEdit()
                    dv(e.RowIndex).Item("charConfigValue") = str1
                    dv(e.RowIndex).EndEdit()
                    Me.dgvGlobal.Refresh()
                    'Me.dgvGlobal.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = str1
                    Me.dgvGlobal.EndEdit()
                Else
                    e.Cancel = True
                    MsgBox("Entry must be TRUE or FALSE", MsgBoxStyle.Information, "Invalid entry...")
                    dgv.CurrentCell = dgv.Rows(e.RowIndex).Cells(e.ColumnIndex)

                End If
            End If
        End If

    End Sub

    Private Sub dgvGlobal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvGlobal.Click

        Dim dgv As DataGridView
        Dim boolDir As Boolean
        Dim boolFile As Boolean
        Dim dv1 As system.data.dataview
        Dim int1 As Short
        Dim var1

        dgv = Me.dgvGlobal
        dv1 = dgv.DataSource
        int1 = dgv.CurrentRow.Index

        'determine if cmd browse needs to be displayed
        boolDir = NZ(dv1(int1).Item("BOOLISDIRECTORY"), False)
        boolFile = NZ(dv1(int1).Item("BOOLISFILE"), False)
        If boolDir Or boolFile Then
            Call SetGlobalBrowse()
            Me.cmdBrowseGlobal.Visible = True
        Else
            Me.cmdBrowseGlobal.Visible = False
        End If

        Dim str1 As String
        str1 = dv1(int1).Item("CHARCONFIGTITLE")
        If StrComp(str1, "Table Date Format", CompareMethod.Text) = 0 Or StrComp(str1, "Text Date Format", CompareMethod.Text) = 0 Then
            str1 = "a"

            Dim strF As String
            Dim tbl2 As System.Data.DataTable
            Dim rows2() As DataRow
            tbl2 = tblDateFormats
            VAR1 = dgv("CHARCONFIGVALUE", int1).Value
            strF = "CHARFORMAT = '" & VAR1 & "'"
            rows2 = tbl2.Select(strF, "INTORDER ASC")
            If rows2.Length = 0 Then
            Else
                str1 = NZ(rows2(0).Item("CHARDESCRIPTION"), "")
                dgv("Example", int1).Value = str1 & " for Sep 1, 2006"
            End If
        ElseIf StrComp(str1, "Default Incurred Sample %Diff Calculation", CompareMethod.Text) = 0 Or StrComp(str1, "Text Date Format", CompareMethod.Text) = 0 Then
            str1 = "a"
            VAR1 = dgv("CHARCONFIGVALUE", int1).Value
            If StrComp(VAR1, "%Difference", CompareMethod.Text) = 0 Then
                str1 = "(Incurred - Original)/Original * 100"
            ElseIf StrComp(VAR1, "Mean %Difference", CompareMethod.Text) = 0 Then
                str1 = "(Incurred - Original)/Mean * 100"
            End If

            dgv("Example", int1).Value = str1


        End If

        'determine if example column should be visible
        'No, don't hide these
        'var1 = NZ(dgv("Example", int1).Value, "")
        'If Len(var1) = 0 Then
        '    dgv.Columns("Example").Visible = False
        'Else
        '    dgv.Columns("Example").Visible = True
        'End If



    End Sub

    Private Sub cmdOrderDropdownbox_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdOrderDropdownbox.Click
        Call OrderDGV(Me.dgvDropdownboxContents, "intOrder", "ID_TBLDROPDOWNBOXCONTENT")
        ' OrderDGV(ByVal dgv As DataGridView, ByVal strS As String, ByVal strID As String)
    End Sub


    Private Sub dgvDropdownboxTitle_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvDropdownboxTitle.MouseEnter
        Me.dgvDropdownboxTitle.Focus()
    End Sub

    Private Sub dgvDropdownboxTitle_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvDropdownboxTitle.SelectionChanged
        If boolFormLoad Then
            Exit Sub
        End If
        Call DropDownsConfigure()

    End Sub

    Private Sub dgvDropdownboxContents_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvDropdownboxContents.CellBeginEdit
        If boolFormLoad Then
            Exit Sub
        End If
        If Me.cmdEdit.Enabled Then
            Exit Sub
        End If

        Dim var1
        Dim dgv As DataGridView

        dgv = Me.dgvDropdownboxContents
        var1 = NZ(dgv.Rows(e.RowIndex).Cells("charValue").Value, "")
        If StrComp(var1, "[None]", CompareMethod.Text) = 0 Then
            e.Cancel = True
            MsgBox("Sorry. This row cannot be edited.", MsgBoxStyle.Information, "Cannot edit...")
        End If

    End Sub

    Private Sub dgvTemplates_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvTemplates.SelectionChanged
        If boolFormLoad Then
            Exit Sub
        End If
        If boolSave Then
            Exit Sub
        End If

        Call TemplatesAttributesConfigure()

    End Sub

    Private Sub dgvUserAttributes_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvUserAttributes.SelectionChanged

        Dim str1 As String

        If boolFormLoad Or boolAddAcct Or boolAddUser Or boolSave Then
            Exit Sub
        End If

        If boolHold Then
            Exit Sub
        End If

        Dim boolHoldT As Boolean
        boolHoldT = boolHold
        boolHold = True 'for cbxPermissions Group

        str1 = Me.ActiveControl.Name
        If StrComp(str1, "dgvUserAttributes", CompareMethod.Text) = 0 Then
        Else
            GoTo end2
        End If

        Dim r, c

        If Me.dgvUserAttributes.CurrentCell Is Nothing Then
            GoTo end2
        Else
            c = Me.dgvUserAttributes.CurrentCell.ColumnIndex
        End If
        If Me.dgvUserAttributes.RowCount = 0 Then
            GoTo end2
        Else
            If Me.dgvUserAttributes.CurrentRow Is Nothing Then
                r = 0
            Else
                r = Me.dgvUserAttributes.CurrentCell.RowIndex
            End If
        End If

        'If r = rowUID Then
        '    GoTo end1
        'Else
        '    rowUID = r
        'End If

        Dim intRow As Short
        Dim dgv As DataGridView
        Dim var1
        dgv = Me.dgvUserAttributes
        'intRow = dgv.CurrentRow.Index
        If dgv.Rows.Count = 0 Then
            GoTo end1
        ElseIf dgv.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = dgv.CurrentRow.Index
        End If
        var1 = dgv("charUserID", intRow).Value
        If Len(NZ(var1, "")) = 0 Then
        Else
            'Call ConfigureAccountPermissions()
            Call ConfigureAccountPermissionsGroup()
        End If

        Call EvaluateUserAccounts()

        Call DVToPasswordCheckboxValues()

        'dgv.CurrentCell = dgv.Rows(intRow).Cells("CHARUSERID")

end1:

        Me.dgvUserAttributes.CurrentCell = Me.dgvUserAttributes.Rows(r).Cells(c)

        Me.dgvUserAttributes.Refresh()

end2:

        boolHold = boolHoldT

    End Sub

    Sub SetWA()

        Dim intID As Int64
        Dim idUA As Int64
        Dim dgv As DataGridView = Me.dgvUserAttributes
        Dim intRow As Int16
        Dim strUID As String
        Dim var1
        Dim boolCanc As Boolean = False

        If dgv.CurrentCell Is Nothing Then
            boolCanc = True
            Me.CHARLDAP.Clear()
            Me.CHARNETWORKACCOUNT.Clear()
            If Me.cbxWatsonAccount.Items.Count = 0 Then
            Else
                Me.cbxWatsonAccount.SelectedIndex = -1
            End If
            GoTo end1
        End If

        intRow = dgv.CurrentRow.Index
        idUA = dgv("ID_TBLUSERACCOUNTS", intRow).Value
        Dim dtbl As DataTable = tblUserAccounts
        Dim strF As String
        strF = "ID_TBLUSERACCOUNTS = " & idUA
        Dim rows() As DataRow = dtbl.Select(strF)

        If boolAccess Then
            GoTo end2
        End If

        Try

            If dgv.CurrentCell Is Nothing Then
                intID = 0
            Else

                If rows.Length = 0 Then
                    intID = 0
                Else
                    intID = rows(0).Item("ID_TBLWATSONACCOUNT")
                End If
            End If

            Me.ID_TBLWATSONACCOUNT.Text = intID

            If intID = 0 Then
                Me.cbxWatsonAccount.SelectedIndex = -1
            Else
                'find userid

                If boolAccess Then
                Else
                    Dim dtbl1 As DataTable = tblWatsonUsers
                    Dim str1 As String = Me.cbxWatsonAccount.Text

                    strF = "USERID = " & intID

                    'strF = "LOGINNAME = '" & str1 & "'"
                    Dim rows1() As DataRow = dtbl1.Select(strF)
                    If rows1.Length = 0 Then
                        Me.cbxWatsonAccount.SelectedIndex = -1
                    Else
                        strUID = rows1(0).Item("LOGINNAME")
                    End If

                    'find struid in cbx
                    Try
                        Me.cbxWatsonAccount.SelectedIndex = Me.cbxWatsonAccount.FindString(strUID)
                    Catch ex As Exception
                        var1 = ex.Message
                    End Try


                End If

            End If

        Catch ex As Exception
            var1 = ex.Message
        End Try

end2:


        'now do LDAP stuff
        If boolCanc Then
            Me.CHARLDAP.Clear()
            Me.CHARNETWORKACCOUNT.Clear()
        Else
            var1 = NZ(rows(0).Item("CHARLDAP"), "")
            Me.CHARLDAP.Text = var1

            var1 = NZ(rows(0).Item("CHARNETWORKACCOUNT"), "")
            Me.CHARNETWORKACCOUNT.Text = var1
        End If



end1:


    End Sub

    Private Sub dgvUsers_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvUsers.SelectionChanged

        Call dgvUserSelectionChange()

    End Sub

    Sub dgvUserSelectionChange()

        Dim dv As System.Data.DataView
        Dim intRow As Short
        Dim var1

        If boolFormLoad Or boolAddUser Or boolSave Then
            Exit Sub
        End If

        Dim r, c
        Try
            c = Me.dgvUsers.CurrentCell.ColumnIndex
        Catch ex As Exception

        End Try

        If Me.dgvUsers.RowCount = 0 Then
            GoTo end2
        Else
            If Me.dgvUsers.CurrentRow Is Nothing Then
                r = 0
            Else
                r = Me.dgvUsers.CurrentCell.RowIndex
            End If
        End If

        If r = rowUA Then
            GoTo end1
        Else
            rowUA = r
        End If

end1:

        Call ConfigureUserAccountAttributes(False)

        Call EvaluateUserAccounts()

        Call ShowAccounts(0)

        Call FillUserboolA(Me.dgvUserAttributes)

        Call UserIDActions()

        Try
            Me.dgvUsers.CurrentCell = Me.dgvUsers.Rows(r).Cells(c)

        Catch ex As Exception

        End Try

        Me.dgvUsers.Refresh()

end2:


    End Sub

    Private Sub cmdResetGlobal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdResetGlobal.Click
        Call DoCancelGlobalTab()
    End Sub

    Private Sub cmdAddDropdownbox_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAddDropdownbox.Click
        Dim tbl As System.Data.DataTable
        Dim dv As system.data.dataview
        Dim dgv As DataGridView
        Dim intRow As Short
        Dim introws As Short
        Dim var1, var2

        Dim tblMax As System.Data.DataTable
        Dim rowsMax() As DataRow
        Dim strFMax As String
        Dim maxID


        maxID = GetMaxID("tblDropdownBoxContent", 1, True)
        '20190219 LEE: Don't need anymore. Used GetMaxID
        'Call PutMaxID("tblDropdownBoxContent", maxID)

        intRow = Me.dgvDropdownboxTitle.CurrentRow.Index
        var1 = Me.dgvDropdownboxTitle.Item("id_tblDropdownboxName", intRow).Value

        dgv = Me.dgvDropdownboxContents
        introws = dgv.RowCount

        tbl = tblDropdownBoxContent
        Dim nr As DataRow = tbl.NewRow
        nr.BeginEdit()
        nr("id_tblDropdownBoxContent") = maxID
        nr("id_tblDropdownboxName") = var1
        nr("intOrder") = introws + 1
        nr.EndEdit()
        tbl.Rows.Add(nr)

        Call DropDownsConfigure()

        Dim Count1 As Short
        'For Count1 = 0 To introws - 1
        '    dgv.Rows(Count1).Cells("charValue").Selected = False
        'Next
        dgv.ClearSelection()
        dgv.CurrentCell = dgv.Rows(introws).Cells("charValue")
        'dgv.Rows(introws).Cells("charValue").Selected = True
        dgv.BeginEdit(True)


    End Sub

    Private Sub dgvHooks_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvHooks.CellContentClick
        'fill boolActive column
        Dim str1 As String
        Dim dgv As DataGridView
        Dim var1
        Dim bool As Boolean
        Dim dv As system.data.dataview

        If e.RowIndex < 0 Then
            Exit Sub
        End If

        dgv = Me.dgvHooks

        dv = dgv.DataSource
        str1 = dgv.Columns(e.ColumnIndex).Name
        If StrComp(str1, "boolA", CompareMethod.Text) = 0 Then
            dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)

            var1 = dgv.Rows(e.RowIndex).Cells("boolA").Value
            If IsDBNull(var1) Then
            Else
                dv(e.RowIndex).BeginEdit()
                If var1 Then
                    dv(e.RowIndex).Item("BOOLINCLUDE") = -1
                Else
                    dv(e.RowIndex).Item("BOOLINCLUDE") = 0
                End If
                dv(e.RowIndex).EndEdit()
            End If
        ElseIf StrComp(str1, "boolS", CompareMethod.Text) = 0 Then
            dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)

            var1 = dgv.Rows(e.RowIndex).Cells("boolS").Value
            If IsDBNull(var1) Then
            Else
                dv(e.RowIndex).BeginEdit()
                If var1 Then
                    dv(e.RowIndex).Item("BOOLSHOW") = -1
                Else
                    dv(e.RowIndex).Item("BOOLSHOW") = 0
                End If
                dv(e.RowIndex).EndEdit()
            End If

        End If

    End Sub

    Private Sub cmdAddHook_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAddHook.Click
        Dim dtbl As System.Data.DataTable
        Dim dv As system.data.dataview
        Dim dv1 As system.data.dataview
        Dim int1 As Int32
        Dim strF As String
        Dim Count1 As Short
        Dim varID
        Dim tblMax As System.Data.DataTable
        Dim rowsMax() As DataRow
        Dim strFMax As String
        Dim maxID

        boolAddRow = True

        maxID = GetMaxID("tblHooks", 1, True)
        '20190219 LEE: Don't need anymore. Used GetMaxID
        'Call PutMaxID("tblHooks", maxID)

        dtbl = tblHooks
        Dim nr As DataRow = dtbl.NewRow

        Dim boolF As Boolean = boolFormLoad

        boolFormLoad = True
        nr.BeginEdit()
        nr.Item("id_tblHooks") = maxID
        nr.Item("BOOLINCLUDE") = -1 'True
        nr.Item("BOOLSHOW") = 0 'False
        nr.Item("BOOLERROR") = 0 'False
        nr.Item("boolA") = True
        nr.Item("boolS") = False
        'nr.Item("numincr") = incr1
        nr.EndEdit()
        dtbl.Rows.Add(nr)

        boolFormLoad = boolF


    End Sub

    Private Sub dgvHooks_CellValidated(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvHooks.CellValidated
        'fill boolActive column
        Dim str1 As String
        Dim dgv As DataGridView
        Dim var1
        Dim bool As Boolean
        Dim dv As system.data.dataview

        dgv = Me.dgvHooks

        dv = dgv.DataSource
        str1 = dgv.Columns(e.ColumnIndex).Name
        If StrComp(str1, "boolA", CompareMethod.Text) = 0 Then
            dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)

            var1 = dgv.Rows(e.RowIndex).Cells("boolA").Value
            If IsDBNull(var1) Then
            Else
                dv(e.RowIndex).BeginEdit()
                If var1 Then
                    dv(e.RowIndex).Item("BOOLINCLUDE") = -1
                Else
                    dv(e.RowIndex).Item("BOOLINCLUDE") = 0
                End If
                dv(e.RowIndex).EndEdit()
            End If
        End If

    End Sub

    Private Sub dgvHooks_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgvHooks.CellValidating
        'cells cannot be null
        Dim dgv As DataGridView
        Dim str1 As String
        Dim str2 As String
        Dim bool As Boolean

        dgv = Me.dgvHooks
        str1 = dgv.Columns(e.ColumnIndex).Name
        bool = False
        Select Case str1
            Case "id_tblTab1"
                bool = True
            Case "CHARHOOK"
                bool = True
            Case "CHARCONNECTIONSTRING"
                bool = True
        End Select
        If bool Then 'continue
            If Len(NZ(e.FormattedValue, "")) = 0 Then
                e.Cancel = True
                str2 = "This cell cannot be blank."
                MsgBox(str2, MsgBoxStyle.Information, "Invalid entry...")
            End If
        End If


    End Sub

    Private Sub cmdResetHooks_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdResetHooks.Click
        Call DoCancelHooksTab()
    End Sub

    Private Sub dgvGlobal_CurrentCellDirtyStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvGlobal.CurrentCellDirtyStateChanged
        If boolFormLoad Then
            Exit Sub
        End If
        If Me.cmdEdit.Enabled Then
            Exit Sub
        End If
        If boolDirty Then
            boolDirty = False
            Exit Sub
        End If


        boolDirty = True

        'Me.dgvGlobal.CommitEdit(DataGridViewDataErrorContexts.Commit)
        Dim strL As String
        Dim dgv As DataGridView
        Dim int1 As Short
        Dim int2 As Short
        Dim var1
        Dim strF As String
        Dim str1 As String
        Dim intRow As Short
        Dim intCol As Short
        Dim dv1 As system.data.dataview


        int1 = Me.lbxGlobal.SelectedIndex
        strL = Me.lbxGlobal.Items(int1)

        'if strL = global then evaluate date format
        If StrComp(strL, "Global Settings", CompareMethod.Text) = 0 Then
            Dim boolDir As Boolean
            dgv = Me.dgvGlobal
            dv1 = dgv.DataSource
            int1 = dgv.CurrentRow.Index

            str1 = dv1(int1).Item("CHARCONFIGTITLE")
            If StrComp(str1, "Table Date Format", CompareMethod.Text) = 0 Or StrComp(str1, "Text Date Format", CompareMethod.Text) = 0 Then

                Me.dgvGlobal.CommitEdit(DataGridViewDataErrorContexts.Commit)

                str1 = "a"
                'Dim VAR1
                'Dim strF As String
                Dim tbl2 As System.Data.DataTable
                Dim rows2() As DataRow
                tbl2 = tblDateFormats
                var1 = dgv("CHARCONFIGVALUE", int1).Value
                strF = "CHARFORMAT = '" & var1 & "'"
                rows2 = tbl2.Select(strF, "INTORDER ASC")
                If rows2.Length = 0 Then
                Else
                    str1 = NZ(rows2(0).Item("CHARDESCRIPTION"), "")
                    dgv("Example", int1).Value = str1 & " for Sep 1, 2006"
                End If
                dgv.AutoResizeColumns()
            ElseIf StrComp(str1, "Default Incurred Sample %Diff Calculation", CompareMethod.Text) = 0 Or StrComp(str1, "Text Date Format", CompareMethod.Text) = 0 Then

                Me.dgvGlobal.CommitEdit(DataGridViewDataErrorContexts.Commit)

                'Dim VAR1
                str1 = "a"
                var1 = dgv("CHARCONFIGVALUE", int1).Value
                If StrComp(var1, "%Difference", CompareMethod.Text) = 0 Then
                    str1 = "(Incurred - Original)/Original * 100"
                ElseIf StrComp(var1, "Mean %Difference", CompareMethod.Text) = 0 Then
                    str1 = "(Incurred - Original)/Mean * 100"
                End If

                dgv("Example", int1).Value = str1
                dgv.AutoResizeColumns()
            End If

        End If

    End Sub

    Private Sub dgvNickNames_CurrentCellDirtyStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvNickNames.CurrentCellDirtyStateChanged
        If boolFormLoad Then
            Exit Sub
        End If
        If Me.cmdEdit.Enabled Then
            Exit Sub
        End If
        If boolDirty Then
            boolDirty = False
            Exit Sub
        End If

        boolDirty = True

        Dim strL As String
        Dim dgv As DataGridView
        Dim intRow As Short
        Dim intCol As Short
        Dim dv1 As system.data.dataview
        Dim bool As Boolean

        dgv = Me.dgvNickNames

        'if strL = global then evaluate date format
        intRow = dgv.CurrentRow.Index
        intCol = dgv.CurrentCell.ColumnIndex
        strL = dgv.Columns(intCol).Name
        If StrComp(strL, "boolI", CompareMethod.Text) = 0 Then
            dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
            dv1 = dgv.DataSource
            'fill with correct bool Value
            bool = dgv.Rows(intRow).Cells(intCol).Value
            dv1(intRow).BeginEdit()
            If bool Then
                dv1(intRow).Item("boolInclude") = -1
            Else
                dv1(intRow).Item("boolInclude") = 0
            End If
            dv1(intRow).EndEdit()
        End If


    End Sub

    Private Sub dgvCorporateAddresses_CurrentCellDirtyStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvCorporateAddresses.CurrentCellDirtyStateChanged
        If boolFormLoad Then
            Exit Sub
        End If
        If Me.cmdEdit.Enabled Then
            Exit Sub
        End If
        If boolDirty Then
            boolDirty = False
            Exit Sub
        End If

        boolDirty = True

        Dim strL As String
        Dim dgv As DataGridView
        Dim intRow As Short
        Dim intCol As Short
        Dim dv1 As system.data.dataview
        Dim bool As Boolean

        dgv = Me.dgvCorporateAddresses

        'if strL = global then evaluate date format
        intRow = dgv.CurrentRow.Index
        intCol = dgv.CurrentCell.ColumnIndex
        strL = dgv.Columns(intCol).Name
        If StrComp(strL, "boolI", CompareMethod.Text) = 0 Then
            dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
            dv1 = dgv.DataSource
            'fill with correct bool Value
            bool = dgv.Rows(intRow).Cells(intCol).Value
            dv1(intRow).BeginEdit()
            If bool Then
                dv1(intRow).Item("BOOLINCLUDEINTITLE") = -1
            Else
                dv1(intRow).Item("BOOLINCLUDEINTITLE") = 0
            End If
            dv1(intRow).EndEdit()
        End If

    End Sub

    Private Sub cmdRefreshHook_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdRefreshHook.Click
        Dim intRow As Short
        Dim dgv As DataGridView
        Dim str1 As String

        Cursor.Current = Cursors.WaitCursor

        dgv = Me.dgvHooks
        If dgv.RowCount = 0 Then
            Exit Sub
        ElseIf dgv.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = dgv.CurrentRow.Index
        End If

        str1 = NZ(dgv("CHARHOOK", intRow).Value, "")
        If Len(str1) = 0 Then
            Exit Sub
        Else
            Select Case str1
                Case "CRLWor_AnalRefStandard"
                    Call HookFill_CRL_AnalRefStandard()
                    Call ComboBoxCRLAnalRefFill()

                    If boolHook1 Then
                        MsgBox("Successfull refresh", MsgBoxStyle.Information, "Success!")
                    End If
            End Select
            Call HookAnalysis()
        End If

        Cursor.Current = Cursors.Default

    End Sub

    Private Sub dgvCorporateAddresses_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvCorporateAddresses.MouseEnter
        Me.dgvCorporateAddresses.Focus()
    End Sub

    Private Sub dgvDropdownboxContents_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgvDropdownboxContents.CellValidating

        Dim intCol As Short
        Dim dgv As DataGridView
        Dim strCol As String
        Dim var1
        Dim boolE As Boolean
        Dim strM As String

        strM = "Number must be integer"

        dgv = Me.dgvDropdownboxContents
        intCol = e.ColumnIndex
        strCol = dgv.Columns(intCol).Name
        boolE = False
        If StrComp(strCol, "INTORDER", CompareMethod.Text) = 0 Then
            'number must be integer
            var1 = e.FormattedValue
            If IsNumeric(var1) Then
                If IsInt(var1) Then

                Else
                    boolE = True
                End If
            Else
                boolE = True
            End If
        End If

        If boolE Then
            e.Cancel = True
            MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
        End If

    End Sub

    Private Sub dgvDropdownboxContents_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvDropdownboxContents.MouseEnter
        Me.dgvDropdownboxContents.Focus()
    End Sub

    Private Sub dgvGlobal_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvGlobal.MouseEnter
        Me.dgvGlobal.Focus()
    End Sub

    Private Sub dgvHooks_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvHooks.MouseEnter
        Me.dgvHooks.Focus()
    End Sub

    Private Sub dgvNickNames_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvNickNames.MouseEnter
        Me.dgvNickNames.Focus()
    End Sub

    Private Sub dgvPermissions_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.dgvPermissions.Focus()
    End Sub

    Private Sub dgvTemplateAttributes_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvTemplateAttributes.MouseEnter
        Me.dgvTemplateAttributes.Focus()
    End Sub

    Private Sub cbxModules_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxModules.SelectedIndexChanged

        Dim tbl As System.Data.DataTable
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim row() As DataRow
        Dim ct1 As Short
        Dim Count1 As Short
        Dim intForm As Short
        Dim bool As Boolean

        If Me.chkEditMode.Checked Then
            bool = True
        Else
            bool = False
        End If

        tbl = tblTab1

        str3 = Me.cbxModules.SelectedItem
        Select Case str3
            Case "StudyDoc Administration"
                intForm = 4
                str1 = "LABIntegrity StudyDoc" & ChrW(8482) & " - StudyDoc Administration"
            Case "Report Writer"
                intForm = 2
                str1 = "LABIntegrity StudyDoc" & ChrW(8482) & " - Report Writer Administration"
            Case "Study Designer"
                intForm = -1
                str1 = "LABIntegrity StudyDoc" & ChrW(8482) & " - Study Designer Administration"
        End Select

        str2 = GetVersion()
        str1 = str1 & " v" & str2 & gUserLabel ' " - User: Guest"
        'str2 = system.windows.forms.application.Info.Description

        Me.Text = str1

        str1 = "intForm = " & intForm
        str2 = "intOrder ASC"
        row = tbl.Select(str1, str2)

        'fill lbxTab1
        ct1 = row.Length
        Me.lbxTab1.Items.Clear()
        For Count1 = 0 To ct1 - 1
            str1 = row(Count1).Item("charItem")
            Me.lbxTab1.Items.Add(str1)
        Next

        'select first item in lbxtab1
        If Me.lbxTab1.Items.Count = 0 Then
        Else
            Me.lbxTab1.SelectedItem = 0
        End If

        Select Case str3
            Case "StudyDoc Administration"
                Call PasswordInitialize()
                Call GlobalConfigure()

            Case "Report Writer"
                Call GlobalInitialize()
                Call GlobalConfigure()

                'check to ensure user is allowed
                Dim boolA As Boolean
                boolA = BOOLADMINISTRATION
                If boolA = 0 Then
                    Dim strM As String
                    str1 = Me.cbxModules.SelectedItem
                    If StrComp(str1, "Report Writer", CompareMethod.Text) = 0 Then
                        'strM = "User does not have permission to edit the Report Writer Administration window."
                        'MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
                        Me.cmdEdit.Enabled = False
                        'Me.lblRestricted.Visible = True

                    End If
                End If

        End Select

        'select first item in lbxtab1
        If Me.lbxTab1.Items.Count = 0 Then
        Else
            Dim boolT As Boolean = boolFormLoad
            boolFormLoad = True
            Me.lbxTab1.SelectedItem = 0
            boolFormLoad = boolT
            Call lbxTab1Change()
        End If

        If intForm = -1 Then
            Call HideAllPages()
            Me.cmdEdit.Enabled = False
            str1 = "Under Construction..."
            MsgBox(str1, MsgBoxStyle.Information, str1)
        Else
            Me.cmdEdit.Enabled = True
        End If

        Call EvaluateUserAccounts()

        'check for admin permissions
        Select Case str3
            Case "StudyDoc Administration"

                Dim boolA As Boolean
                boolA = BOOLADMINISTRATIONADMIN ' rows(0).Item("BOOLADMINISTRATIONADMIN")
                If boolA = 0 Then

                    Dim strM As String

                    strM = "User may view, but does not have permission to edit the StudyDoc Administration window."
                    MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
                    Me.cmdEdit.Enabled = False
                    'Me.lblRestricted.Visible = True

                    'str1 = Me.cbxModules.SelectedItem
                    'If StrComp(str1, "Report Writer", CompareMethod.Text) = 0 Then
                    '    strM = "User may view, but does not have permission to edit the Report Writer Administration window."
                    '    MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
                    '    Me.cmdEdit.Enabled = False
                    '    'Me.lblRestricted.Visible = True

                    'End If
                End If

            Case "Report Writer"


                'check to ensure user is allowed
                Dim boolA As Boolean
                boolA = BOOLADMINISTRATION
                If boolA = 0 Then

                    Dim strM As String
                    strM = "User may view, but does not have permission to edit the Report Writer Administration window."
                    MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
                    Me.cmdEdit.Enabled = False
                    'Me.lblRestricted.Visible = True

                End If

        End Select
    End Sub

    Private Sub frmAdministration_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed

        Call DoExit()

    End Sub

    Private Sub lbx1_DrawItem(sender As Object, e As DrawItemEventArgs) Handles lbx1.DrawItem

        'https://social.msdn.microsoft.com/Forums/en-US/aa2ba97a-5e93-4e7d-ab06-f7919939092a/listbox-items-line-spacing?forum=Vsexpressvb

        'e.Graphics.DrawString(lbxTab1.Items(e.Index).ToString, lbxTab1.Font, Brushes.Black, e.Bounds.Left, ((e.Bounds.Height - lbxTab1.Font.Height) \ 2) + e.Bounds.Top)

        Dim var1
        Try
            Dim drawBrush As New SolidBrush(Me.lbx1.ForeColor)
            e.Graphics.DrawString(lbx1.Items(e.Index).ToString, lbx1.Font, drawBrush, e.Bounds.Left, ((e.Bounds.Height - lbx1.Font.Height) \ 2) + e.Bounds.Top)
        Catch ex As Exception
            var1 = ex.Message
        End Try

    End Sub

    Private Sub lbx1_MeasureItem(sender As Object, e As MeasureItemEventArgs) Handles lbx1.MeasureItem


        'https://social.msdn.microsoft.com/Forums/en-US/aa2ba97a-5e93-4e7d-ab06-f7919939092a/listbox-items-line-spacing?forum=Vsexpressvb

        'itemheight at the default font settings is 20
        e.ItemHeight = 22

    End Sub

    Private Sub lbxGlobal_DrawItem(sender As Object, e As DrawItemEventArgs) Handles lbxGlobal.DrawItem

        'https://social.msdn.microsoft.com/Forums/en-US/aa2ba97a-5e93-4e7d-ab06-f7919939092a/listbox-items-line-spacing?forum=Vsexpressvb

        'e.Graphics.DrawString(lbxTab1.Items(e.Index).ToString, lbxTab1.Font, Brushes.Black, e.Bounds.Left, ((e.Bounds.Height - lbxTab1.Font.Height) \ 2) + e.Bounds.Top)

        Dim var1
        Try
            Dim drawBrush As New SolidBrush(Me.lbxGlobal.ForeColor)
            e.Graphics.DrawString(lbxGlobal.Items(e.Index).ToString, lbxGlobal.Font, drawBrush, e.Bounds.Left, ((e.Bounds.Height - lbxGlobal.Font.Height) \ 2) + e.Bounds.Top)
        Catch ex As Exception
            var1 = ex.Message
        End Try

    End Sub

    Private Sub lbxGlobal_MeasureItem(sender As Object, e As MeasureItemEventArgs) Handles lbxGlobal.MeasureItem

        'https://social.msdn.microsoft.com/Forums/en-US/aa2ba97a-5e93-4e7d-ab06-f7919939092a/listbox-items-line-spacing?forum=Vsexpressvb

        'itemheight at the default font settings is 20
        e.ItemHeight = 22

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

    Sub cbxWatsonAccountFill()

        Dim var1

        If boolAccess Then
        Else

            Try
                Me.cbxWatsonAccount.DataSource = tblWatsonUsers
                Me.cbxWatsonAccount.DisplayMember = "LOGINNAME"
                'Me.cbxWatsonAccount.DataBindings.Add(New Binding("USERID", ds, "customers.CustToOrders.OrderDate"))
                Me.cbxWatsonAccount.DataBindings.Add(New Binding("USERID", dsDoPr, "TBLWATSONUSERS.USERID"))

            Catch ex As Exception
                var1 = ex.Message
            End Try

            Me.cbxWatsonAccount.SelectedIndex = -1

        End If

    End Sub

    Private Sub frmAdministration_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'https://social.msdn.microsoft.com/Forums/en-US/aa2ba97a-5e93-4e7d-ab06-f7919939092a/listbox-items-line-spacing?forum=Vsexpressvb
        '20170729 LEE: tried to add some padding between listbox items, but solution deactivates all other properties of the listbox, like selected-background-color
        'could probably do some more digging to address problem, but too much work for little benefit

        'Me.lbx1.DrawMode = DrawMode.OwnerDrawVariable
        'Me.lbxGlobal.DrawMode = DrawMode.OwnerDrawVariable
        'Me.lbxTab1.DrawMode = DrawMode.OwnerDrawVariable

        Call modExtensionMethods.DoubleBufferedControl(Me, True)
        Call DoubleBufferControl(Me, "dgv")
        Call DoubleBufferControl(Me, "gb")
        Call DoubleBufferControl(Me, "lv")

        Call ControlDefaults(Me)

        Call HideAllPages()

        Call SetLDAP()


        Dim str1 As String
        str1 = "LABIntegrity StudyDoc" & ChrW(8482) & " Administration"
        Me.Text = str1

        ReDim arrPswd(10, 100)
        ctPswd = 0

        boolFormLoad = True

        Me.lblOpen.Dock = DockStyle.Fill
        Me.lblOpen.Visible = True

        Me.lblOpen.Refresh()

        Me.chkEditMode.Checked = False

        Call Filllbx1()

        Call FilllbxPerm()

        Call cbxWatsonAccountFill()


        Call dgvPermissionsConfigure()

        Call FillcbxPermissionsGroup(False)

        Call FilllvPermissions() 'this also calls configurelvpermissions

        Call FormLoad()

        boolFormLoad = False

        'Call ShowlvPermissions()

        Call ConfigMOFandRFC()

        Call DoEnableCompliance()

        Call EvaluateUserAccounts()

        Call FillCompliance()

        Call FinalFrozen()

        'Call LockAll(True)

        boolFormLoad = False

        'pesky
        Call ConfigureUserAccountAttributes(False)

        Call EvaluateUserAccounts()

        Call FillUserboolA(Me.dgvUserAttributes)

        Call UserIDActions()

        boolESigTemp = gboolESig
        boolAuditTrailTemp = gboolAuditTrail

        Me.lblOpen.Visible = False

        'pesky
        Call LockAll(True)
        Me.dgvUsers.CurrentCell = Me.dgvUsers.Item("CHARLASTNAME", 1)
        Me.dgvUsers.Rows(1).Selected = True
        Call ConfigureUserAccountAttributes(False)
        Call frmAdministration_ToolTipSet()

        'place cmdSymbol
        Dim a, b, c, d
        Dim bw As Int16 = (Me.Width - Me.ClientSize.Width) / 2 'form border width
        Dim tbh As Int16 = Me.Height - Me.ClientSize.Height - 2 * bw 'titlebar height

        a = Me.pan1.Left + Me.pan1.Width ' Me.Width - bw
        b = Me.cmdSymbol.Width
        c = a - b
        Me.cmdSymbol.Left = c

        'select 2nd row in user table
        Try
            Me.dgvUsers.CurrentCell = Me.dgvUsers.Item("CHARLASTNAME", 1)

        Catch ex As Exception

        End Try

    End Sub

    Sub FillcbxPermissionsGroup(boolReset As Boolean)

        Dim dtbl As DataTable
        Dim strF As String

        dtbl = tblPermissions

        Dim rows() As DataRow
        rows = tblPermissions.Select("ID_TBLPERMISSIONS > 0", "CHARPERMISSIONSNAME")

        If boolReset Then
            Dim id As Int64
            Dim str1 As String



            id = Me.cbxPermissionsGroup.SelectedValue
            str1 = Me.cbxPermissionsGroup.Text
            Me.cbxPermissionsGroup.DataSource = rows
            Me.cbxPermissionsGroup.DisplayMember = "CHARPERMISSIONSNAME"
            Me.cbxPermissionsGroup.ValueMember = "ID_TBLPERMISSIONS"
            'put it back to id
            Me.cbxPermissionsGroup.SelectedIndex = Me.cbxPermissionsGroup.FindStringExact(str1)

        Else
            Me.cbxPermissionsGroup.DataSource = rows
            Me.cbxPermissionsGroup.DisplayMember = "CHARPERMISSIONSNAME"
            Me.cbxPermissionsGroup.ValueMember = "ID_TBLPERMISSIONS"
        End If




    End Sub

    Private Sub cbxModulesPers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)


        Call EvaluateUserAccounts()


    End Sub

    Sub SelectPerms(boolS As Boolean)

        Dim int1 As Short
        Dim Count1 As Short

        Dim str1 As String
        Dim lv As ListView
        Dim boolDo As Boolean

        str1 = Me.lbx1.SelectedItem
        boolDo = True
        Select Case str1
            Case "StudyDoc Administration"
                lv = Me.lvPermissionsAdmin
            Case "Report Writer"
                lv = Me.lvPermissions
            Case "Report Templates"
                lv = Me.lvPermissionsReportTemplate
            Case "Final Reports"
                lv = Me.lvPermissionsFinalReport
            Case "Study Designer"
                boolDo = False
        End Select

        If boolDo Then
            int1 = lv.Items.Count
            For Count1 = 0 To int1 - 1
                lv.Items(Count1).Checked = boolS
            Next
        End If

    End Sub

    Private Sub cmdSelectAllPermissions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSelectAllPermissions.Click

        Call SelectPerms(True)


    End Sub

    Private Sub cmdDeselectAllPermissions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDeselectAllPermissions.Click

        Call SelectPerms(False)

    End Sub

    Sub SetFontSize()


    End Sub

    Private Sub cmdIncreaseFont_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdIncreaseFont.Click

        Call DoFont(True)

    End Sub

    Private Sub cmdDecreaseFont_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDecreaseFont.Click

        Call DoFont(False)
        
    End Sub


    Sub DoFont(boolI As Boolean)

        Dim dgv As DataGridView
        Dim ctrl As Control
        Dim str1 As String
        Dim str2 As String
        Dim ctrlp As Control
        Dim str3 As String
        Dim str4 As String
        Dim fs As Single
        Dim fn As String

        Dim allC As New List(Of Control)

        Call FindAllControlRecursive(allC, Me)

        For Each ctrlA As Control In allC

            str1 = ctrlA.Name
            str2 = Mid(str1, 1, 3)
            If StrComp(str2, "dgv", CompareMethod.Text) = 0 Then
                dgv = ctrlA
                fs = dgv.DefaultCellStyle.Font.Size
                If boolI Then
                    fs = fs + 1
                Else
                    fs = fs - 1
                End If

                fn = dgv.DefaultCellStyle.Font.Name
                'dgv.DefaultCellStyle.Font = New Font(fn, fs)
                dgv.DefaultCellStyle.Font = New System.Drawing.Font(fn, fs)
                dgv.AutoResizeColumns()
            End If

        Next

    End Sub

    Sub DoEnableCompliance()

        If Me.rbAuditTrailOn.Checked Then
            Me.gbESig.Enabled = True
            Me.panRFCOptions.Enabled = True
        Else
            Me.gbESig.Enabled = False
            Me.panRFCOptions.Enabled = False
        End If

        If Me.rbESigOn.Checked Then
            Me.panESigOptions.Enabled = True
        Else
            Me.panESigOptions.Enabled = False
        End If

    End Sub


    Sub ConfigMOFandRFC()

        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView

        dgv1 = Me.dgvMOS
        dgv2 = Me.dgvRFC

        Dim dtbl1 As System.Data.DataTable = tblMeaningOfSig
        Dim dtbl2 As System.Data.DataTable = tblReasonForChange

        'add non-bound columns
        If dtbl1.Columns.Contains("DEFAULTCHK") Then
        Else
            Dim col1 As New DataColumn
            col1.ColumnName = "DEFAULTCHK"
            col1.Caption = "Default"
            col1.DataType = System.Type.GetType("System.Boolean")
            dtbl1.Columns.Add(col1)
        End If

        'add non-bound columns
        If dtbl2.Columns.Contains("DEFAULTCHK") Then
        Else
            Dim col2 As New DataColumn
            col2.ColumnName = "DEFAULTCHK"
            col2.Caption = "Default"
            col2.DataType = System.Type.GetType("System.Boolean")
            dtbl2.Columns.Add(col2)
        End If

        'fill new columns
        Call SetDefaultChecks()

        Dim dv1 As System.Data.DataView = New DataView(dtbl1)
        Dim dv2 As System.Data.DataView = New DataView(dtbl2)

        dv1.AllowDelete = False
        dv1.AllowEdit = True
        dv1.AllowNew = False
        dv1.Sort = "INTORDER ASC"

        dv2.AllowDelete = False
        dv2.AllowEdit = True
        dv2.AllowNew = False
        dv2.Sort = "INTORDER ASC"

        dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgv2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        dgv1.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dgv2.DefaultCellStyle.WrapMode = DataGridViewTriState.True

        dgv1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        dgv2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

        dgv1.DataSource = dv1
        dgv2.DataSource = dv2

        dgv1.Columns("ID_TBLMEANINGOFSIG").Visible = False
        dgv1.Columns("CHARMEANINGOFSIG").Visible = True
        dgv1.Columns("CHARMEANINGOFSIG").HeaderText = "Meaning of Signature"
        dgv1.Columns("INTORDER").Visible = True
        dgv1.Columns("INTORDER").HeaderText = "Order"
        dgv1.Columns("INTORDER").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        dgv1.Columns("BOOLINCLUDE").Visible = False
        dgv1.Columns("BOOLDEFAULT").Visible = False
        dgv1.Columns("BOOLDEFAULT").HeaderText = "Default"
        dgv1.Columns("DEFAULTCHK").Visible = True
        dgv1.Columns("DEFAULTCHK").HeaderText = "Default"

        dgv1.Columns("CHARMEANINGOFSIG").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        dgv1.Columns("INTORDER").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
        dgv1.Columns("DEFAULTCHK").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader

        dgv1.Columns("CHARMEANINGOFSIG").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        dgv1.Columns("INTORDER").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv1.Columns("DEFAULTCHK").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        dgv1.Columns("CHARMEANINGOFSIG").SortMode = DataGridViewColumnSortMode.NotSortable
        dgv1.Columns("INTORDER").SortMode = DataGridViewColumnSortMode.NotSortable
        dgv1.Columns("DEFAULTCHK").SortMode = DataGridViewColumnSortMode.NotSortable


        dgv1.RowHeadersWidth = 25


        dgv2.Columns("ID_TBLREASONFORCHANGE").Visible = False
        dgv2.Columns("CHARREASONFORCHANGE").Visible = True
        dgv2.Columns("CHARREASONFORCHANGE").HeaderText = "Reason For Change"
        dgv2.Columns("INTORDER").Visible = True
        dgv2.Columns("INTORDER").HeaderText = "Order"
        dgv2.Columns("INTORDER").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        dgv2.Columns("BOOLINCLUDE").Visible = False
        dgv2.Columns("BOOLDEFAULT").Visible = False
        dgv2.Columns("BOOLDEFAULT").HeaderText = "Default"
        dgv2.Columns("DEFAULTCHK").Visible = True
        dgv2.Columns("DEFAULTCHK").HeaderText = "Default"

        dgv2.Columns("CHARREASONFORCHANGE").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        dgv2.Columns("INTORDER").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
        dgv2.Columns("DEFAULTCHK").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader

        dgv2.Columns("CHARREASONFORCHANGE").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        dgv2.Columns("INTORDER").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv2.Columns("DEFAULTCHK").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        dgv2.Columns("CHARREASONFORCHANGE").SortMode = DataGridViewColumnSortMode.NotSortable
        dgv2.Columns("INTORDER").SortMode = DataGridViewColumnSortMode.NotSortable
        dgv2.Columns("DEFAULTCHK").SortMode = DataGridViewColumnSortMode.NotSortable


        dgv2.RowHeadersWidth = 25


    End Sub

    Private Sub rbAuditTrailOff_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbAuditTrailOff.CheckedChanged

        Call DoEnableCompliance()

        'Call FillCompliance() 'do this so that test form shows properly

    End Sub

    Private Sub rbESigOff_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbESigOff.CheckedChanged

        Call DoEnableCompliance()

        'Call FillCompliance() 'do this so that test form shows properly


    End Sub

    Private Sub dgvRFC_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgvRFC.CellValidating

        If boolFormLoad Then
            Exit Sub
        End If

        Dim str1 As String
        Dim str2 As String
        Dim dgv1 As DataGridView = Me.dgvRFC
        Dim strE As String
        Dim id As Int64
        Dim id1 As Int16
        Dim var1, var2
        Dim Count1 As Short
        Dim boolC As Boolean

        str1 = dgv1.Columns(e.ColumnIndex).Name

        Select Case str1

            Case "CHARREASONFORCHANGE"

                'make sure there are no duplicates
                Dim dv As System.Data.DataView = dgv1.DataSource
                Dim tbl As System.Data.DataTable = dv.ToTable
                Dim rows() As System.Data.DataRow
                Dim strF As String

                str2 = e.FormattedValue
                str2 = Replace(str2, "'", "_", 1, -1, CompareMethod.Text)
                str2 = Replace(str2, """", "_", 1, -1, CompareMethod.Text)

                If Len(str2) = 0 Then
                    strE = "Entry cannot be blank."
                    MsgBox(strE, MsgBoxStyle.Information, "Invalid entry...")
                    e.Cancel = True
                Else

                    id = dgv1("ID_TBLREASONFORCHANGE", e.RowIndex).Value
                    strF = "CHARREASONFORCHANGE = '" & str2 & "'" ' AND ID_TBLREASONFORCHANGE = " & id
                    rows = tbl.Select(strF)
                    If rows.Length = 0 Then
                    Else
                        id1 = rows(0).Item("ID_TBLREASONFORCHANGE")
                        If id1 = id Then
                        Else
                            strE = "The entry:" & ChrW(10) & ChrW(10) & ChrW(9) & str2 & ChrW(10) & ChrW(10) & "is a duplicate entry and not allowed."
                            MsgBox(strE, MsgBoxStyle.Information, "Invalid entry...")
                            e.Cancel = True
                        End If
                    End If

                End If

            Case "DEFAULTCHK"

                var1 = e.FormattedValue
                id = dgv1("ID_TBLREASONFORCHANGE", e.RowIndex).Value
                boolC = True

                'first check if there are any other defaults
                If var1 Then

                    For Count1 = 0 To dgv1.RowCount - 1
                        id1 = dgv1("ID_TBLREASONFORCHANGE", Count1).Value
                        If id1 = id Then 'ignore
                        Else 'evaluate

                            var1 = NZ(dgv1("DEFAULTCHK", Count1).Value, False)
                            If var1 Then
                                strE = "There is already a Default selected in Reason for Change." & ChrW(10) & ChrW(10) & "The other Default must first be unchecked."
                                MsgBox(strE, MsgBoxStyle.Information, "Invalid entry...")
                                e.Cancel = True
                                boolC = False
                                Exit For
                            End If
                        End If

                    Next

                End If

                var1 = e.FormattedValue
                If var1 Then
                    var2 = -1
                Else
                    var2 = 0
                End If
                dgv1("BOOLDEFAULT", e.RowIndex).Value = var2

        End Select

    End Sub

    Sub SetDefaultChecks()

        Dim dtbl As System.Data.DataTable
        Dim var1
        Dim Count1 As Short
        Dim Count2 As Short

        For Count2 = 1 To 2

            Select Case Count2
                Case 1
                    dtbl = tblReasonForChange
                Case 2
                    dtbl = tblMeaningOfSig
            End Select

            For Count1 = 0 To dtbl.Rows.Count - 1
                var1 = NZ(dtbl.Rows(Count1).Item("BOOLDEFAULT"), 0)
                dtbl.Rows(Count1).BeginEdit()
                If var1 = 0 Then
                    dtbl.Rows(Count1).Item("DEFAULTCHK") = False
                Else
                    dtbl.Rows(Count1).Item("DEFAULTCHK") = True
                End If
                dtbl.Rows(Count1).EndEdit()
            Next

        Next

    End Sub

    Sub FillCompliance()

        Dim dtbl As System.Data.DataTable = tblConfigCompliance
        Dim var1

        var1 = dtbl.Rows(0).Item("BOOLAUDITTRAIL")
        If var1 = 0 Then
            Me.rbAuditTrailOn.Checked = False
            Me.rbAuditTrailOff.Checked = True
        Else
            Me.rbAuditTrailOn.Checked = True
            Me.rbAuditTrailOff.Checked = False
        End If

        var1 = dtbl.Rows(0).Item("BOOLESIG")
        If var1 = 0 Then
            Me.rbESigOn.Checked = False
            Me.rbESigOff.Checked = True
        Else
            Me.rbESigOn.Checked = True
            Me.rbESigOff.Checked = False
        End If

        var1 = dtbl.Rows(0).Item("BOOLLOGGEDONUSER")
        If var1 = 0 Then
            Me.rbOnlyLoggedOn.Checked = False
            Me.rbUserIDChoice.Checked = True
        Else
            Me.rbOnlyLoggedOn.Checked = True
            Me.rbUserIDChoice.Checked = False
        End If

        var1 = dtbl.Rows(0).Item("BOOLMEANINGOFSIG")
        If var1 = 0 Then
            Me.chkMeaningOfSign.Checked = False
        Else
            Me.chkMeaningOfSign.Checked = True
        End If

        var1 = dtbl.Rows(0).Item("BOOLRESTRICTSIG")
        If var1 = 0 Then
            Me.chkSigFreeForm.Checked = False
        Else
            Me.chkSigFreeForm.Checked = True
        End If

        var1 = dtbl.Rows(0).Item("BOOLREASONFORCHANGE")
        If var1 = 0 Then
            Me.chkReasonForChange.Checked = False
        Else
            Me.chkReasonForChange.Checked = True
        End If

        var1 = dtbl.Rows(0).Item("BOOLRESTRICTREASON")
        If var1 = 0 Then
            Me.chkReasonFreeForm.Checked = False
        Else
            Me.chkReasonFreeForm.Checked = True
        End If

    End Sub

    Sub SaveAdminFC()


        Call FillAuditTrailTemp(tblFieldCodes)

        'must add new or deleted tblFieldCode entries to tblCustomFieldCodes
        'order seems to be important here
        Call AddFieldCodes()


        If boolGuWuOracle Then
            Try
                ta_tblFieldCodes.Update(tblFieldCodes)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLFIELDCODES.Merge('ds2005Acc.TBLFIELDCODES, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_tblFieldCodesAcc.Update(tblFieldCodes)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLFIELDCODES.Merge('ds2005Acc.TBLFIELDCODES, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_tblFieldCodesSQLServer.Update(tblFieldCodes)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLFIELDCODES.Merge('ds2005Acc.TBLFIELDCODES, True)
            End Try
        End If

    End Sub

    Sub AddFieldCodes()

        Dim dtbl As System.Data.DataTable = tblFieldCodes
        Dim rows() As DataRow
        Dim rowsFC() As DataRow
        Dim strF As String
        Dim strS As String
        Dim Count1 As Short
        Dim Count2 As Short
        Dim intRows As Short

        Dim tblS As System.Data.DataTable = tblStudies
        Dim tblCFC As System.Data.DataTable = tblCustomFieldCodes
        Dim strFC As String
        Dim idFC As Int64
        Dim idS As Int64
        Dim idCFC As Int64

        Dim intOrder As Int32

        Dim CountA As Short

        For CountA = 1 To 2

            strF = "BOOLCUSTOM = -1"
            strS = "ID_TBLFIELDCODES ASC"

            Select Case CountA
                Case 1 'Added
                    rows = dtbl.Select(strF, strS, DataViewRowState.Added)
                Case 2 'Deleted
                    rows = dtbl.Select(strF, strS, DataViewRowState.Deleted)
            End Select

            intRows = rows.Length

            If intRows = 0 Then
                GoTo nextcounta
            End If

            Select Case CountA
                Case 1 'Added
                    idCFC = GetMaxID("TBLCUSTOMFIELDCODES", (tblS.Rows.Count * intRows), True)
                    '20190219 LEE: Don't need anymore. Used GetMaxID
                    'Call PutMaxID("TBLCUSTOMFIELDCODES", idCFC + (tblS.Rows.Count * intRows) + 1)

                    For Count1 = 0 To intRows - 1
                        strFC = NZ(rows(Count1).Item("CHARFIELDCODE"), "")
                        idFC = rows(Count1).Item("ID_TBLFIELDCODES")
                        For Count2 = 0 To tblS.Rows.Count - 1


                            idCFC = idCFC + 1
                            idS = tblS.Rows(Count2).Item("ID_TBLSTUDIES")

                            'find number of rows
                            strF = "ID_TBLSTUDIES = " & idS
                            Dim rowsCF() As DataRow = tblCustomFieldCodes.Select(strF)
                            intOrder = rowsCF.Length + 1

                            Dim nr As DataRow = tblCFC.NewRow

                            nr.BeginEdit()

                            nr.Item("ID_TBLCUSTOMFIELDCODES") = idCFC
                            nr.Item("ID_TBLSTUDIES") = idS
                            nr.Item("ID_TBLFIELDCODES") = idFC
                            nr.Item("INTORDER") = intOrder

                            nr.Item("BOOLINCLUDE") = 0

                            nr.EndEdit()

                            tblCFC.Rows.Add(nr)

                        Next

                    Next
                Case 2 'Deleted
                    For Count1 = 0 To intRows - 1

                        idFC = rows(Count1).Item("ID_TBLFIELDCODES", DataRowVersion.Original)
                        strF = "ID_TBLFIELDCODES = " & idFC
                        Dim rowsCFC() As DataRow = tblCFC.Select(strF)

                        For Count2 = 0 To rowsCFC.Length - 1
                            rowsCFC(Count2).BeginEdit()
                            rowsCFC(Count2).Delete()
                            rowsCFC(Count2).EndEdit()
                        Next

                    Next
            End Select

nextCountA:

        Next CountA


end1:


        If boolGuWuOracle Then
            Try
                ta_tblCustomFieldCodes.Update(tblCustomFieldCodes)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLCUSTOMFIELDCODES.Merge('ds2005Acc.TBLCUSTOMFIELDCODES, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_tblCustomFieldCodesAcc.Update(tblCustomFieldCodes)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLCUSTOMFIELDCODES.Merge('ds2005Acc.TBLCUSTOMFIELDCODES, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_tblCustomFieldCodesSQLServer.Update(tblCustomFieldCodes)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLCUSTOMFIELDCODES.Merge('ds2005Acc.TBLCUSTOMFIELDCODES, True)
            End Try
        End If

        'set dgvFC in frmh
        Call FillFCRW()



    End Sub

    Private Sub dgvMOS_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvMOS.CellContentClick

        Dim var1, var2, var3
        Dim dgv As DataGridView
        Dim str1 As String
        Dim Count1 As Short
        Dim boolC As Boolean
        Dim strE As String
        Dim id As Int64
        Dim id1 As Int64


        dgv = Me.dgvMOS

        str1 = dgv.Columns(e.ColumnIndex).Name

        Select Case str1
            Case "DEFAULTCHK"

                'find booldefault value
                'var1 = NZ(dgv("BOOLDEFAULT", e.RowIndex).Value, 0)
                'If var1 = 0 Then 'user is attempting to make positive
                '    'check for other Default checks
                '    id = dgv("ID_TBLMEANINGOFSIG", e.RowIndex).Value
                '    For Count1 = 0 To dgv.RowCount - 1
                '        id1 = dgv("ID_TBLMEANINGOFSIG", Count1).Value
                '        If id1 = id Then 'ignore
                '        Else 'evaluate
                '            var1 = dgv("DEFAULTCHK", Count1).Value
                '            If var1 Then
                '                strE = "Click: There is already a Default selected." & ChrW(10) & ChrW(10) & "The other Default must first be unchecked."
                '                MsgBox(strE, MsgBoxStyle.Information, "Invalid entry...")
                '                boolC = False
                '                Exit For
                '            End If
                '        End If
                '    Next
                'End If

                'If boolC Then
                'Else
                '    dgv("BOOLDEFAULT", e.RowIndex).Value = False
                'End If


        End Select


    End Sub

    Private Sub dgvMOS_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgvMOS.CellValidating

        If boolFormLoad Then
            Exit Sub
        End If

        Dim str1 As String
        Dim str2 As String
        Dim dgv1 As DataGridView = Me.dgvMOS
        Dim strE As String
        Dim id As Int64
        Dim id1 As Int64
        Dim var1, var2
        Dim Count1 As Short
        Dim boolC As Boolean

        str1 = dgv1.Columns(e.ColumnIndex).Name

        Select Case str1
            Case "CHARMEANINGOFSIG"

                'make sure there are no duplicates
                Dim dv As System.Data.DataView = dgv1.DataSource
                Dim tbl As System.Data.DataTable = dv.ToTable
                Dim rows() As System.Data.DataRow
                Dim strF As String

                str2 = e.FormattedValue
                str2 = Replace(str2, "'", "_", 1, -1, CompareMethod.Text)
                str2 = Replace(str2, """", "_", 1, -1, CompareMethod.Text)

                If Len(str2) = 0 Then
                    strE = "Entry cannot be blank."
                    MsgBox(strE, MsgBoxStyle.Information, "Invalid entry...")
                    e.Cancel = True
                Else
                    id = dgv1("ID_TBLMEANINGOFSIG", e.RowIndex).Value
                    strF = "CHARMEANINGOFSIG = '" & str2 & "'"
                    rows = tbl.Select(strF)

                    If rows.Length = 0 Then
                    Else
                        id1 = rows(0).Item("ID_TBLMEANINGOFSIG")
                        If id1 = id Then
                        Else
                            strE = "The entry:" & ChrW(10) & ChrW(10) & ChrW(9) & str2 & ChrW(10) & ChrW(10) & "is a duplicate entry and not allowed."
                            MsgBox(strE, MsgBoxStyle.Information, "Invalid entry...")
                            e.Cancel = True
                        End If

                    End If
                End If

            Case "DEFAULTCHK"

                var1 = e.FormattedValue
                id = dgv1("ID_TBLMEANINGOFSIG", e.RowIndex).Value
                boolC = True

                'first check if there are any other defaults

                If var1 Then
                    For Count1 = 0 To dgv1.RowCount - 1
                        id1 = dgv1("ID_TBLMEANINGOFSIG", Count1).Value
                        If id1 = id Then 'ignore
                        Else 'evaluate
                            var1 = dgv1("DEFAULTCHK", Count1).Value
                            If var1 Then
                                strE = "Validating: There is already a Default selected." & ChrW(10) & ChrW(10) & "The other Default must first be unchecked."
                                MsgBox(strE, MsgBoxStyle.Information, "Invalid entry...")
                                e.Cancel = True
                                boolC = False
                                Exit For
                            End If
                        End If
                    Next
                End If

                If boolC Then
                    var1 = e.FormattedValue
                    If var1 Then
                        var2 = -1
                    Else
                        var2 = 0
                    End If
                    dgv1("BOOLDEFAULT", e.RowIndex).Value = var2
                End If

        End Select

    End Sub

    Private Sub dgvRFC_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvRFC.DataError

        If boolFormLoad Then
            Exit Sub
        End If

        Dim dgv1 As DataGridView = Me.dgvRFC
        Dim strE As String
        Dim str1 As String

        str1 = dgv1.Columns(e.ColumnIndex).Name

        Select Case str1
            Case "INTORDER"
                strE = "Entry must be integer"
                MsgBox(strE, MsgBoxStyle.Information, "Invalid entry...")
                e.Cancel = True
        End Select

    End Sub

    Private Sub dgvMOS_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvMOS.DataError

        If boolFormLoad Then
            Exit Sub
        End If

        Dim dgv1 As DataGridView = Me.dgvMOS
        Dim strE As String
        Dim str1 As String

        str1 = dgv1.Columns(e.ColumnIndex).Name

        Select Case str1
            Case "INTORDER"
                strE = "Entry must be integer"
                MsgBox(strE, MsgBoxStyle.Information, "Invalid entry...")
                e.Cancel = True
        End Select

    End Sub

    Private Sub cmdAddMOS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddMOS.Click

        'add a row to underlying dv
        Dim dtbl As System.Data.DataTable = tblMeaningOfSig
        Dim dgv As DataGridView = Me.dgvMOS
        'get last id
        Dim id As Int64
        id = dtbl.Rows(dtbl.Rows.Count - 1).Item("ID_TBLMEANINGOFSIG")
        id = id + 1
        Dim intO As Short = dtbl.Rows(dtbl.Rows.Count - 1).Item("INTORDER")
        intO = intO + 1
        Dim nrow As DataRow = dtbl.NewRow
        nrow.BeginEdit()
        nrow.Item("ID_TBLMEANINGOFSIG") = id
        nrow.Item("INTORDER") = intO
        nrow.Item("BOOLINCLUDE") = -1
        nrow.EndEdit()
        dtbl.Rows.Add(nrow)

        dgv.CurrentCell = dgv.Rows(dtbl.Rows.Count - 1).Cells("CHARMEANINGOFSIG")


    End Sub

    Private Sub cmdAddRFC_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAddRFC.Click

        'add a row to underlying dv
        Dim dtbl As System.Data.DataTable = tblReasonForChange
        'get last id
        Dim id As Int64
        id = dtbl.Rows(dtbl.Rows.Count - 1).Item("ID_TBLREASONFORCHANGE", DataRowVersion.Original)
        id = id + 1
        Dim intO As Short = dtbl.Rows(dtbl.Rows.Count - 1).Item("INTORDER", DataRowVersion.Original)
        intO = intO + 1
        Dim nrow As DataRow = dtbl.NewRow
        nrow.BeginEdit()
        nrow.Item("ID_TBLREASONFORCHANGE") = id
        nrow.Item("INTORDER") = intO
        nrow.Item("BOOLINCLUDE") = -1
        nrow.EndEdit()
        dtbl.Rows.Add(nrow)

        'Me.dgvRFC.CurrentCell = Me.dgvRFC.Rows(dtbl.Rows.Count - 1).Cells("CHARREASONFORCHANGE")
        Try
            Me.dgvRFC.CurrentCell = Me.dgvRFC.Rows(dtbl.Rows.Count - 1).Cells("CHARREASONFORCHANGE")
        Catch ex As Exception
            Try
                Me.dgvRFC.CurrentCell = Me.dgvRFC.Rows(0).Cells("CHARREASONFORCHANGE")
            Catch ex1 As Exception

            End Try
        End Try

    End Sub

    Private Sub cmdRemoveMOS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemoveMOS.Click

        Dim dtbl As System.Data.DataTable = tblMeaningOfSig
        Dim intRows As Short
        Dim dgv As DataGridView = Me.dgvMOS
        Dim Count1 As Short
        Dim int1 As Short
        Dim id As Int64
        Dim rows() As DataRow
        Dim strF As String

        intRows = dgv.RowCount
        Dim arrS(intRows)

        int1 = 0
        For Count1 = 0 To intRows - 1
            If dgv.Rows(Count1).Selected Then
                int1 = int1 + 1
                id = dgv("ID_TBLMEANINGOFSIG", Count1).Value
                arrS(int1) = id
            End If
        Next

        If int1 = 0 Then
            'probably a cell rather than a row is selected
            'select the row
            If dgv.RowCount = 0 Then
            ElseIf dgv.CurrentRow Is Nothing Then
            Else
                int1 = 1
                id = dgv("ID_TBLMEANINGOFSIG", dgv.CurrentRow.Index).Value
                arrS(int1) = id
            End If
        End If

        'delete those rows
        For Count1 = 1 To int1
            id = arrS(Count1)
            strF = "ID_TBLMEANINGOFSIG = " & id
            rows = dtbl.Select(strF)
            rows(0).BeginEdit()
            rows(0).Delete()
            rows(0).EndEdit()
        Next

        If dgv.RowCount = 0 Or int1 = 0 Then
        Else
            'dgv.CurrentCell = dgv.Rows(0).Cells("CHARMEANINGOFSIG")
            Try
                dgv.CurrentCell = dgv.Rows(0).Cells("CHARMEANINGOFSIG")
            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub cmdRemoveRFC_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdRemoveRFC.Click

        Dim dtbl As System.Data.DataTable = tblReasonForChange
        Dim intRows As Short
        Dim dgv As DataGridView = Me.dgvRFC
        Dim Count1 As Short
        Dim int1 As Short
        Dim id As Int64
        Dim rows() As DataRow
        Dim strF As String

        intRows = dgv.RowCount
        Dim arrS(intRows)

        int1 = 0
        For Count1 = 0 To intRows - 1
            If dgv.Rows(Count1).Selected Then
                int1 = int1 + 1
                id = dgv("ID_TBLREASONFORCHANGE", Count1).Value
                arrS(int1) = id
            End If
        Next

        If int1 = 0 Then
            'probably a cell rather than a row is selected
            'select the row
            If dgv.RowCount = 0 Then
            ElseIf dgv.CurrentRow Is Nothing Then
            Else
                int1 = 1
                id = dgv("ID_TBLREASONFORCHANGE", dgv.CurrentRow.Index).Value
                arrS(int1) = id
            End If
        End If

        'delete those rows
        For Count1 = 1 To int1
            id = arrS(Count1)
            strF = "ID_TBLREASONFORCHANGE = " & id
            rows = dtbl.Select(strF)
            rows(0).BeginEdit()
            rows(0).Delete()
            rows(0).EndEdit()
        Next

        If dgv.RowCount = 0 Or int1 = 0 Then
        Else
            Try
                dgv.CurrentCell = dgv.Rows(0).Cells("CHARREASONFORCHANGE")
            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub cmdTestESig_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTestESig.Click

        Dim frm As New frmESig
        frm.boolTest = True
        frm.chkRFC.Checked = Me.chkReasonForChange.Checked
        frm.chkRRFC.Checked = Me.chkReasonFreeForm.Checked
        frm.chkMOS.Checked = Me.chkMeaningOfSign.Checked
        frm.chkRMOS.Checked = Me.chkSigFreeForm.Checked
        frm.rbESigOn.Checked = Me.rbESigOn.Checked

        frm.cmdOK.Enabled = False

        'showdialog will call frm.PlaceC again
        frm.ShowDialog()

        frm.Dispose()

        Exit Sub

        'If Me.rbESigOff.Checked Then
        '    Dim strM As String
        '    strM = "ESig turned off"
        '    MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
        'Else
        '    'Call SaveCompliance(True)

        '    Dim frm As New frmESig

        '    frm.cmdOK.Enabled = False

        '    frm.ShowDialog()

        'End If


    End Sub

    Sub ShowMOS()

        If Me.chkMeaningOfSign.Checked Then
            Me.chkSigFreeForm.Visible = True
        Else
            Me.chkSigFreeForm.Visible = False
        End If

    End Sub

    Sub ShowRFC()

        If Me.chkReasonForChange.Checked Then
            Me.chkReasonFreeForm.Visible = True
        Else
            Me.chkReasonFreeForm.Visible = False
        End If

    End Sub

    Private Sub chkReasonForChange_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkReasonForChange.CheckedChanged

        'Call SaveCompliance() 'do this so that test form shows properly
        Call ShowRFC()

    End Sub

    Private Sub chkReasonFreeForm_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkReasonFreeForm.CheckedChanged

        'Call SaveCompliance() 'do this so that test form shows properly

    End Sub

    Private Sub rbUserIDChoice_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbUserIDChoice.CheckedChanged

        'Call SaveCompliance() 'do this so that test form shows properly

    End Sub

    Private Sub chkMeaningOfSign_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkMeaningOfSign.CheckedChanged

        'Call SaveCompliance() 'do this so that test form shows properly

        Call ShowMOS()

    End Sub

    Private Sub chkSigFreeForm_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSigFreeForm.CheckedChanged

        'Call SaveCompliance() 'do this so that test form shows properly

    End Sub

    Private Sub rbAuditTrailOn_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbAuditTrailOn.CheckedChanged

    End Sub

    Private Sub dgvFC_CellValidated(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvFC.CellValidated

        Dim dgv As DataGridView
        Dim intRow As Short
        Dim intCol As Short
        Dim strCName As String
        Dim strF As String
        Dim strS As String
        Dim strA As String
        Dim strM As String
        Dim str1 As String
        Dim boolE As Boolean = False
        Dim boolU As Boolean = True

        dgv = Me.dgvFC

        intRow = e.RowIndex
        intCol = e.ColumnIndex
        strCName = dgv.Columns(intCol).Name

        If StrComp(strCName, "CHARFIELDCODE", CompareMethod.Text) = 0 Then

            strA = NZ(dgv(intCol, intRow).Value, "")

            If Len(strA) = 0 Then
                strM = "Entry cannot be blank"
                GoTo end1
            End If

            'look for brackets
            str1 = Mid(strA, 1, 1)
            If StrComp(str1, "[", CompareMethod.Text) = 0 Then
            Else
                strA = "[" & strA
                boolU = True
            End If
            str1 = Mid(strA, Len(strA), 1)
            If StrComp(str1, "]", CompareMethod.Text) = 0 Then
            Else
                strA = strA & "]"
                boolU = True
            End If

            If boolU Then
                dgv(intCol, intRow).Value = strA
                dgv.CommitEdit(DataGridViewDataErrorContexts.CurrentCellChange)
            End If

        End If

end1:

    End Sub

    Private Sub dgvFC_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgvFC.CellValidating

        Dim dgv As DataGridView
        Dim intRow As Short
        Dim intCol As Short
        Dim strCName As String
        Dim strF As String
        Dim strS As String
        Dim strA As String
        Dim strM As String
        Dim str1 As String
        Dim boolE As Boolean = False
        Dim boolU As Boolean = True

        If Me.cmdEdit.Enabled Then
            Exit Sub
        End If

        If Me.cmdEdit.Enabled = False And Me.cmdSave.Enabled = False Then
            Exit Sub
        End If

        dgv = Me.dgvFC

        intRow = e.RowIndex
        intCol = e.ColumnIndex
        strCName = dgv.Columns(intCol).Name

        'only apply to column1
        If intCol <> 1 Then
            boolE = False
            GoTo end1
        End If

        strA = e.FormattedValue

        If Len(strA) = 0 Then
            strM = "Entry cannot be blank"
            boolE = True
            GoTo end1
        End If

        'look for brackets
        str1 = Mid(strA, 1, 1)
        If StrComp(str1, "[", CompareMethod.Text) = 0 Then
        Else
            strA = "[" & strA
            boolU = True
        End If
        str1 = Mid(strA, Len(strA), 1)
        If StrComp(str1, "]", CompareMethod.Text) = 0 Then
        Else
            strA = strA & "]"
            boolU = True
        End If

        'cannot be duplicate
        Dim dv As System.Data.DataView

        dv = dgv.DataSource

        Dim tbl As System.Data.DataTable = dv.ToTable
        strF = "CHARFIELDCODE = '" & strA & "'"
        Dim rows() As DataRow
        ' rows = tbl.Select(strF)
        'Try
        '    rows = tbl.Select(strF)
        'Catch ex As Exception

        'End Try
        rows = tbl.Select(strF)
        If rows.Length <= 1 Then 'account for this entry
        Else
            strM = "The Field Code" & ChrW(10) & ChrW(10) & strA & ChrW(10) & ChrW(10) & "already exists"
            boolE = True
            GoTo end1
        End If

        'check for sigfig value
        Dim boolCancel As Boolean
        Dim strMod As String = "Administration - Custom Field Codes"
        Dim strSource As String = "Field Code Value"
        'now check for column limit
        If boolCLExceeded(strCName, "TBLFIELDCODES", e.FormattedValue, True, strMod, strSource) Then
            e.Cancel = True
            GoTo end1
        End If

        If boolU Then
            dgv(intCol, intRow).Value = strA
            dgv.CommitEdit(DataGridViewDataErrorContexts.CurrentCellChange)
        End If


end1:
        If boolE Then
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            e.Cancel = True
        End If

    End Sub

    Private Sub cmdAddFC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddFC.Click

        Dim dtbl As System.Data.DataTable = tblFieldCodes

        Dim nr As DataRow = dtbl.NewRow
        Dim id As Int64
        Dim id1 As Int64

        id = GetMaxID("TBLFIELDCODES", 1, True)
        'now get maxid from tlbFieldCodes because of programatic additions
        Dim rows() As DataRow
        Dim strF As String
        Dim strS As String

        strF = "ID_TBLFIELDCODES >0"
        strS = "ID_TBLFIELDCODES DESC"
        rows = dtbl.Select(strF, strS)

        id1 = rows(0).Item("ID_TBLFIELDCODES")

        If id1 > id Then
            id = id1 + 1
        Else
            id = id + 1
        End If
        Call PutMaxID("TBLFIELDCODES", id)

        nr.BeginEdit()
        nr("ID_TBLFIELDCODES") = id
        nr("BOOLCUSTOM") = -1
        nr("CHARGROUP") = "Custom"
        nr("CHARFIELDCODE") = "[New]"
        nr.Item("CHAREXAMPLE") = "NA"
        nr.Item("CHARDESCRIPTION") = "NA"
        nr.Item("CHARTAB") = "NA"
        nr.Item("CHARTABLE") = "NA"
        nr.Item("UPSIZE_TS") = Now

        nr.EndEdit()

        dtbl.Rows.Add(nr)

        Try
            'NDL 7-Dec-2015 Added this to put cursor into edit mode in Added row.
            dgvFC.Refresh()
            dgvFC.CurrentCell = dgvFC.Rows(dgvFC.Rows.Count - 1).Cells("CHARFIELDCODE")
            dgvFC.BeginEdit(True)
        Catch ex As Exception

        End Try

    End Sub

    Sub DoFCAdminCancel()

        tblFieldCodes.RejectChanges()
        tblCustomFieldCodes.RejectChanges()

        'Call ResetFieldCodes()

    End Sub

    Private Sub cmdResetFC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdResetFC.Click

        Call DoFCAdminCancel()

    End Sub

    Private Sub dgvFC_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvFC.CellContentClick

    End Sub

    Private Sub cmdRemoveFC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemoveFC.Click

        Dim dtbl As System.Data.DataTable = tblFieldCodes

        Dim dgv As DataGridView = Me.dgvFC
        Dim dRow As DataGridViewRow

        Dim id As Int64
        Dim dv As System.Data.DataView

        'now get maxid from tlbFieldCodes because of programatic additions
        Dim rows() As DataRow
        Dim strF As String
        Dim strS As String
        Dim boolHit As Boolean = False

        Dim arrID(dgv.Rows.Count)
        Dim intNumID As Short = 0

        Dim Count1 As Short
        Dim Count2 As Int32

        For Each dRow In dgv.Rows

            If dRow.Selected Then
                boolHit = True
                intNumID = intNumID + 1
                id = dgv("ID_TBLFIELDCODES", dRow.Index).Value
                arrID(intNumID) = id
                'strF = "ID_TBLFIELDCODES = " & id
                'rows = dtbl.Select(strF)
                'rows(0).BeginEdit()
                'rows(0).Delete()
                'rows(0).EndEdit()
            End If

        Next


        If boolHit Then

            For Count1 = 1 To intNumID
                id = arrID(Count1)
                strF = "ID_TBLFIELDCODES = " & id
                Erase rows
                rows = dtbl.Select(strF)
                rows(0).BeginEdit()
                rows(0).Delete()
                rows(0).EndEdit()
            Next

        Else
            Dim strM As String
            strM = "An entire row(s) must be selected in order to remove it/them"
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
        End If

    End Sub

    Private Sub gbxPassword_Enter(sender As Object, e As EventArgs) Handles gbxPassword.Enter

    End Sub

    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click

    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click

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

    Private Sub dgvUserAttributes_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvUserAttributes.CellContentClick

    End Sub

    Private Sub lbx1_Click(sender As Object, e As System.EventArgs) Handles lbx1.Click

        Call ShowlvPermissions()

    End Sub

    Private Sub cmdAddPM_Click(sender As System.Object, e As System.EventArgs) Handles cmdAddPM.Click

        Dim intR As Short
        Dim strM As String
        Dim boolCont As Boolean = False
        Dim Count1 As Short
        Dim int1 As Short
        Dim int2 As Short
        Dim str1 As String
        Dim str2 As String
        Dim strF As String
        Dim strS As String
        Dim boolHit As Boolean
        Dim maxID As Int64
        Dim dgv As DataGridView = Me.dgvPermissions
        Dim strName As String
        Dim strPermBase As String
        Dim var1

        Dim frm As New frmAddPermGroup

        'fill frm cbx, but leave out blank item
        frm.cbxPermBase.Items.Clear()
        For Count1 = 0 To Me.cbxPermBase.Items.Count - 1
            frm.cbxPermBase.Items.Add(Me.cbxPermBase.Items(Count1).ToString)
        Next
        frm.cbxPermBase.SelectedIndex = 0

        frm.ShowDialog()

        If frm.boolCancel Then
            frm.Dispose()
            GoTo end1
        Else

            'record name
            strName = frm.txtPN.Text

            'record permissions base
            strPermBase = frm.cbxPermBase.SelectedItem.ToString
            frm.Dispose()
        End If


        'add strName to lbxPermGroup
        'do this by adding to tblPermissions
        Dim nr As DataRow = tblPermissions.NewRow
        'get maxid for tblpermissions
        maxID = GetMaxID("TBLPERMISSIONS", 1, True)
        '20190219 LEE: Don't need anymore. Used GetMaxID
        'Call PutMaxID("TBLPERMISSIONS", maxID)

        nr.Item("ID_TBLPERMISSIONS") = maxID
        nr.Item("CHARPERMISSIONSNAME") = strName
        tblPermissions.Rows.Add(nr)

        'first unselect all rows
        For Count1 = 0 To dgv.Rows.Count - 1
            dgv.Rows(Count1).Selected = False
        Next

        'select this item
        Dim intRow As Short
        intRow = dgv.CurrentRow.Index
        For Count1 = 0 To dgv.Rows.Count - 1
            str1 = dgv("CHARPERMISSIONSNAME", Count1).Value
            If StrComp(strName, str1, CompareMethod.Text) = 0 Then
                dgv.Rows(Count1).Selected = True
                intRow = Count1
                Exit For
            End If
        Next

        'apply according to permbase
        Dim rowsS() As DataRow
        Dim rowsD() As DataRow
        strF = "CHARPERMISSIONSNAME = '" & strPermBase & "'"
        rowsS = tblPermissions.Select(strF)
        Dim id As Int64
        id = rowsS(0).Item("ID_TBLPERMISSIONS")

        strF = "ID_TBLPERMISSIONS = " & maxID
        rowsD = tblPermissions.Select(strF)

        Dim nrD As DataRow = rowsD(0)
        nrD.BeginEdit()
        For Count1 = 0 To tblPermissions.Columns.Count - 1

            str1 = tblPermissions.Columns(Count1).ColumnName
            Select Case str1
                Case "ID_TBLPERMISSIONS"
                Case "CHARPERMISSIONSNAME"
                Case "UPSIZE_TS"
                Case Else
                    var1 = rowsS(0).Item(str1)
                    nrD(str1) = var1
            End Select
        Next
        nrD.EndEdit()

        Call ConfigureLVPermissions(maxID)


        'disabled dgv
        dgv.Enabled = False


        Me.lblDo.Visible = True

end1:

    End Sub


    Sub FilllbxPerm()

        Dim strF As String
        Dim strS As String
        Dim intRow As Int32

        strF = "ID_TBLPERMISSIONS > 0"
        strS = "CHARPERMISSIONSNAME ASC"


        Dim rows() As DataRow = tblPermissions.Select(strF, strS)

        'don't bind because need a blank row

        intRow = -1
        If Me.cbxPermBase.Items.Count > 0 Then
            intRow = Me.cbxPermBase.SelectedIndex
        End If

        Me.cbxPermBase.Items.Clear()

        Dim Count1 As Int32
        Dim str1 As String

        For Count1 = 0 To rows.Length - 1
            str1 = rows(Count1).Item("CHARPERMISSIONSNAME")
            Me.cbxPermBase.Items.Add(str1)
        Next

        If intRow = -1 Then
            Me.cbxPermBase.SelectedIndex = 0
        Else
            Me.cbxPermBase.SelectedIndex = intRow
        End If



    End Sub

    Private Sub cbxPermBase_Click(sender As Object, e As System.EventArgs) Handles cbxPermBase.Click



    End Sub

    Private Sub cmdRemovePM_Click(sender As System.Object, e As System.EventArgs) Handles cmdRemovePM.Click

        Dim intR As Short
        Dim strM As String
        Dim strPerm As String
        Dim rows() As DataGridViewSelectedRowCollection
        Dim Count1 As Int32
        Dim Count2 As Int32
        Dim Count3 As Int32
        Dim intRow As Int32 = 0
        Dim idP As Int64
        Dim idUA As Int64
        Dim strF As String
        Dim strF1 As String
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim id1 As Int64
        Dim id2 As Int64
        Dim strUID As String
        Dim strUN As String

        Dim boolGo As Boolean = True


        For Count1 = 0 To Me.dgvPermissions.SelectedRows.Count - 1
            str1 = Me.dgvPermissions.SelectedRows(Count1).Cells("CHARPERMISSIONSNAME").Value
            If StrComp(str1, "Administrator", CompareMethod.Text) = 0 Then
                strM = "The 'Administrator' Permissions Group cannot be deleted."
                MsgBox(strM, vbInformation, "Invalid action...")
                boolGo = False
                Exit For
            Else

            End If
        Next

        If boolGo Then
            strM = "Are you sure you wish to Delete the selected Permissions Group?"
            strM = strM & ChrW(10) & ChrW(10) & "User Accounts will be examined for this Permissions Group membership."
            strM = strM & ChrW(10) & ChrW(10) & "If a User Accounts are found who are members of this group,"
            strM = strM & ChrW(10) & "    those User Accounts will be displayed and the action canceled."
            intR = MsgBox(strM, vbOKCancel, "Are you sure?")
            If intR = 1 Then 'OK
            Else
                GoTo end1
            End If
        Else
            GoTo end1
        End If

        'check to make sure id isn't used by a user account

        Dim dtbl As New DataTable

        Dim col3 As New DataColumn
        col3.ColumnName = "strPerm"
        dtbl.Columns.Add(col3)
        Dim col2 As New DataColumn
        col2.ColumnName = "strUID"
        dtbl.Columns.Add(col2)
        Dim col1 As New DataColumn
        col1.ColumnName = "strUN"
        dtbl.Columns.Add(col1)

        For Count3 = 0 To Me.dgvPermissions.SelectedRows.Count - 1

            idP = Me.dgvPermissions.SelectedRows(Count3).Cells("ID_TBLPERMISSIONS").Value
            strPerm = Me.dgvPermissions.SelectedRows(Count3).Cells("CHARPERMISSIONSNAME").Value

            strF = "ID_TBLPERMISSIONS = " & idP
            Dim rowsUID() As DataRow
            rowsUID = tblUserAccounts.Select(strF)

            If rowsUID.Length = 0 Then
            Else

                Dim rowsPers()
                For Count1 = 0 To rowsUID.Length - 1
                    strUID = rowsUID(Count1).Item("CHARUSERID")
                    id1 = rowsUID(Count1).Item("ID_TBLPERSONNEL")
                    strF1 = "ID_TBLPERSONNEL = " & id1
                    rowsPers = tblPersonnel.Select(strF1)
                    For Count2 = 0 To rowsPers.Length - 1
                        str1 = rowsPers(Count2).ITEM("CHARFIRSTNAME")
                        str2 = NZ(rowsPers(Count2).ITEM("CHARMIDDLENAME"), "")
                        str3 = rowsPers(Count2).ITEM("CHARLASTNAME")
                        str4 = str3 & ", " & str1 & " " & str2
                        strUN = Trim(str4)

                        Dim nr As DataRow = dtbl.NewRow
                        nr.BeginEdit()
                        nr("strUID") = strUID
                        nr("strUN") = strUN
                        nr("strPerm") = strPerm
                        nr.EndEdit()
                        dtbl.Rows.Add(nr)
                    Next
                Next

            End If

        Next


        If dtbl.Rows.Count = 0 Then

            'delete the row
            idP = Me.dgvPermissions.SelectedRows(0).Cells("ID_TBLPERMISSIONS").Value
            Dim rowsP() As DataRow
            strF = "ID_TBLPERMISSIONS = " & idP
            rowsP = tblPermissions.Select(strF)
            rowsP(0).BeginEdit()
            rowsP(0).Delete()
            rowsP(0).EndEdit()

        Else

            Dim frm As New frmPermGroupDelete
            strM = "The selected Permissions Group cannot be deleted."
            strM = strM & ChrW(10) & "The following User IDs and corresponding User Names are assigned to the Permissions Group:"
            frm.lbl1.Text = strM

            strM = "Please re-assign the Permissions Group assigned to these User ID's"
            frm.lbl2.Text = strM

            Dim dgv As DataGridView = frm.dgvPerm

            Dim dv As DataView = New DataView(dtbl)
            dv.AllowNew = False
            dv.AllowEdit = False
            dv.AllowDelete = False
            dv.Sort = "strUID ASC"
            dgv.DataSource = dv
            dgv.Columns("strPerm").HeaderText = "Permissions Group"
            dgv.Columns("strUID").HeaderText = "User ID"
            dgv.Columns("strUN").HeaderText = "User Name"
            With dgv.ColumnHeadersDefaultCellStyle
                .BackColor = Color.Navy
                .ForeColor = Color.White
                .Font = New Font(dgv.Font, FontStyle.Bold)
            End With
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            frm.ShowDialog()

        End If

end1:

    End Sub


    Private Sub lbx1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles lbx1.SelectedIndexChanged

        Call ShowlvPermissions()

    End Sub

    Private Sub dgvPermissions_SelectionChanged(sender As Object, e As System.EventArgs) Handles dgvPermissions.SelectionChanged

        Call PermRowEnter()

    End Sub

    Sub PermRowEnter()

        Try
            'fill according to row
            Dim idP As Int64
            Dim dgv As DataGridView
            Dim intRow As Int16

            dgv = Me.dgvPermissions

            Try
                intRow = dgv.CurrentRow.Index
            Catch ex As Exception
                'MsgBox("1: " & ex.Message)
                GoTo end1
            End Try

            idP = dgv("ID_TBLPERMISSIONS", intRow).Value

            Call ConfigureLVPermissions(idP)

            'MsgBox("Good")

        Catch ex As Exception

            'MsgBox("2: " & ex.Message)

        End Try

end1:

    End Sub

    Private Sub cbxPermBase_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cbxPermBase.SelectedIndexChanged

        If boolFormLoad Then
            Exit Sub
        End If

        Dim intR As Short
        Dim strM As String
        Dim str1 As String

        str1 = Me.cbxPermBase.SelectedText.ToString

        strM = "Do you wish to apply the Permissions Group '" & str1 & "' to the checkbox table below?"
        intR = MsgBox(strM, vbOKCancel, "Continue?")
        If intR = 1 Then 'continue
        Else
            GoTo end1
        End If

end1:

    End Sub

    Private Sub cbxPermissionsGroup_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cbxPermissionsGroup.SelectedIndexChanged

        If IsSelected(Me.cbxPermissionsGroup) Then
        Else
            Exit Sub
        End If

        'enter value in tblUserAccounts
        'get id_tblpermissions


        'useful information
        'http://blogs.msdn.com/b/jaredpar/archive/2006/11/07/combobox-selecteditem-selectedvalue-selectedwhat.aspx

        '        DataBinding Case

        'In this case the properties will have the following values 
        '•SelectedItem - For gets this will return the actual object in the DataSource that is being displayed in the ComboBox.  For sets if the value exists in the DataSource, it will be selected, otherwise the operation will complete without an exception but won't actually do anything.
        '•SelectedValue - This property depends on the value of ValueMember.  
        'If the property ValueMember is not Nothing the ComboBox will look for a member on SelectedItem with the name specified in ValueMember and return that.  This is also the value displayed in the ComboBox.

        'If the property ValueMember is Nothing, then the SelectedValue will return .ToString() on the SelectedItem

        '•SelectedIndex - Index of SelectedItem in the DataSource 

        'Non-DataBinding Case
        '•SelectedItem - Gets and Sets both go to the currently selected object from the Items collection
        '•SelectedValue - Will be Nothing/null 
        '•SelectedIndex - Index in the Items collection of the SelectedItem


        'cbxPermissionsGroup is bound
        'cbxPermissionsGroup.valuemember = id_tblpermissions
        If boolFormLoad Then
            Exit Sub
        End If

        If boolHold Then
            Exit Sub
        End If

        Dim intRow As Int32
        Dim dgv As DataGridView
        Dim idA As Int64
        Dim var1

        Try
            dgv = Me.dgvUserAttributes
            intRow = dgv.CurrentRow.Index
            idA = dgv("ID_TBLUSERACCOUNTS", intRow).Value

            Dim cbx As ComboBox
            cbx = Me.cbxPermissionsGroup
            Dim strPerm As String
            Dim idP As Int64
            Dim dtblUA As DataTable = tblUserAccounts

            'strPerm = cbx.SelectedValue
            'Dim rowsP() As DataRow
            'rowsP = dtblP.Select("CHARPERMISSIONSNAME = '" & strPerm & "'")
            'idP = rowsP(0).Item("ID_TBLPERMISSIONS")

            idP = cbx.SelectedValue

            Dim rowsUA() As DataRow
            rowsUA = dtblUA.Select("ID_TBLUSERACCOUNTS = " & idA)
            Dim nr As DataRow
            nr = rowsUA(0)
            nr.BeginEdit()
            nr("ID_TBLPERMISSIONS") = idP
            nr.EndEdit()
         

        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

    End Sub

    Private Sub lvPermissions_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles lvPermissions.SelectedIndexChanged

    End Sub

    Private Sub dgvUserAttributes_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles dgvUserAttributes.DataError

        'put some text here to eliminate data error prob

    End Sub
    Private Sub frmAdministration_ToolTipSet()

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

        Try

            'Set general buttons
            toolTip1.SetToolTip(Me.cmdEdit, "Change to Editing Mode")
            toolTip1.SetToolTip(Me.cmdSave, "Save all changes")
            toolTip1.SetToolTip(Me.cmdCancel, "Cancel Unsaved Changes")
            toolTip1.SetToolTip(Me.cmdExit, "Exit Report Writer Administration")
            toolTip1.SetToolTip(Me.cmdIncreaseFont, "Increase font size")
            toolTip1.SetToolTip(Me.cmdDecreaseFont, "Reduce font size")

            'REPORT WRITER ADMINISTRATIVE OPTIONS
            'TAB 1: DropDownLists
            toolTip1.SetToolTip(Me.Label1, "Configure pre-defined text options for Dropdowns")
            toolTip1.SetToolTip(Me.cmdOrderDropdownbox, "Re-Order options (after changing the display order #’s)")
            toolTip1.SetToolTip(Me.cmdAddDropdownbox, "Add Row for new dropdown list option")
            toolTip1.SetToolTip(Me.cmdResetDropdownbox, "Undo unsaved changes (this tab only)")
            'Grid
            Me.dgvDropdownboxTitle.Columns.Item("CHARDROPDOWNNAME").ToolTipText = "Choose one of these dropdown lists"
            Me.dgvDropdownboxContents.Columns.Item("CHARVALUE").ToolTipText = "Enter/edit options that appear on dropdown list"

            'TAB 2: Corporate Addresses
            toolTip1.SetToolTip(Me.Label2, "Configure addresses used in the report")
            toolTip1.SetToolTip(Me.cmdResetCorporateAddressses, "Undo unsaved changes (this tab only)")
            toolTip1.SetToolTip(Me.cmdAddCorporateAddress, "Add new corporate address to StudyDoc")
            dgvNickNames.Columns.Item("CHARNICKNAME").ToolTipText = "Choose corporate Address"
            dgvNickNames.Columns.Item("BoolI").ToolTipText = "Make active (if address is current)"
            dgvCorporateAddresses.Columns.Item("BoolI").ToolTipText = "Include this label in the Report Title"

            'TAB 3: Study Template Definitions
            toolTip1.SetToolTip(Me.Label5, "Manage options for StudyDoc templates")
            toolTip1.SetToolTip(Me.cmdAddTemplate, "Create new StudyDoc template from StudyDoc study")
            toolTip1.SetToolTip(Me.cmdResetDefineReports, "Undo unsaved changes (this tab only)")

            dgvTemplates.Columns.Item("CHARTEMPLATENAME").ToolTipText = "Enter name for StudyDoc template"
            dgvTemplates.Columns.Item("StudyName").ToolTipText = "Choose study to base template on"
            dgvTemplates.Columns.Item("boolA").ToolTipText = "Make StudyDoc template active/inactive"

            'TAB 4: Global Parameters
            toolTip1.SetToolTip(Me.lblGlobalParameters, "Set options used in all StudyDoc templates")
            toolTip1.SetToolTip(Me.cmdResetGlobal, "Undo unsaved changes (this tab only)")
            'dgvGlobal

            'TAB 5: Custom Field Codes
            toolTip1.SetToolTip(Me.lblFieldCodes, "Add variables to reference in report (set in Top Level Data)")
            toolTip1.SetToolTip(Me.cmdAddFC, "Add field code row (at end)")
            toolTip1.SetToolTip(Me.cmdRemoveFC, "Remove row (entire row must be selected)")
            toolTip1.SetToolTip(Me.cmdResetFC, "Undo unsaved changes (this tab only)")
            'dgvFC.Columns.Item("").ToolTipText = ""


            'GLOBAL ADMINISTRATIVE OPTIONS
            'Tab 1: User Accounts
            toolTip1.SetToolTip(Me.cbxModules, "Change screen to other module’s Administration options")
            toolTip1.SetToolTip(Me.cmdAddUser, "Add new person to StudyDoc group")
            toolTip1.SetToolTip(Me.cmdAddUserID, "Add a userID for the selected person")
            'User Names
            Me.dgvUsers.Columns.Item("boolA").ToolTipText = "A: De-select for users no longer in StudyDoc group (Users cannot be deleted)"
            Me.dgvUsers.Columns.Item("DTACTIVATED").ToolTipText = "Most recent date the user was activated"
            Me.dgvUsers.Columns.Item("DTDEACTIVATED").ToolTipText = "Most recent date the user was deactivated."
            Me.dgvUsers.Columns.Item("CHARCOMMENTS").ToolTipText = "Add user-specific comments here (optional)"
            Me.dgvUsers.Columns.Item("CHAREMAILADDRESS").ToolTipText = "Add user email here (optional)"
            Me.dgvUserAttributes.Columns.Item("boolA").ToolTipText = "A: De-select for user ID’s no longer in use  (User ID’s cannot be deleted)"
            Me.dgvUserAttributes.Columns.Item("CHARUSERID").ToolTipText = "A user can have multiple ID’s, each with different permissions"
            Me.dgvUserAttributes.Columns.Item("DTACTIVATED").ToolTipText = "Most recent date the user ID was activated."
            Me.dgvUserAttributes.Columns.Item("DTDEACTIVATED").ToolTipText = "Most recent date the user ID was de-activated."
            Me.dgvUserAttributes.Columns.Item("CHARCOMMENTS").ToolTipText = "Comments on this User ID (optional)"
            toolTip1.SetToolTip(Me.chkAccountIsLockedOut, "De-select to unlock StudyDoc for this user ID")
            toolTip1.SetToolTip(Me.cbxPermissionsGroup, "Groups are set in the Administration ""Permissions Manager""")

            'Tab 2: Global Parameters
            Me.dgvGlobal.ShowCellToolTips = True
            ' For Row-based tooltips, see dgvGlobal_CellFormatting below

            'Tab 3: Hooks

            'Tab 4: Compliance Settings
            toolTip1.SetToolTip(Me.cmdAddMOS, "Add to list of available options for ""meaning of signature""")

            toolTip1.SetToolTip(Me.cmdRemoveMOS, "Remove selected item from list of options for ""meaning of signature""")
            toolTip1.SetToolTip(Me.cmdAddRFC, "Add to list of available options for ""reason for change""")
            toolTip1.SetToolTip(Me.cmdRemoveRFC, "Remove selected item from list of options for ""reason for change""")
            toolTip1.SetToolTip(Me.chkSigFreeForm, "Select if user is not allowed to type in freeform meaning of signature")
            toolTip1.SetToolTip(Me.chkReasonFreeForm, "Select if user is not allowed to type in freeform reason for change")
            Me.dgvMOS.Columns.Item("INTORDER").ToolTipText = "Order shown on option list"
            Me.dgvMOS.Columns.Item("DEFAULTCHK").ToolTipText = "Select if this is the preferred default"
            Me.dgvRFC.Columns.Item("INTORDER").ToolTipText = "Order shown on option list"
            Me.dgvRFC.Columns.Item("DEFAULTCHK").ToolTipText = "Select if this is the preferred default"

            'Tab 5: Permissions Manager
            toolTip1.SetToolTip(Me.lblPM, "Set up permissions groups (to assign user ID’s to)")
            toolTip1.SetToolTip(Me.cmdAddPM, "Add new permissions group")
            toolTip1.SetToolTip(Me.cmdRemovePM, "Delete permissions group")
            toolTip1.SetToolTip(Me.cmdRemovePM, "Delete permissions group")
            toolTip1.SetToolTip(Me.cmdSelectAllPermissions, "Select all permissions (make allowable)")
            toolTip1.SetToolTip(Me.cmdDeselectAllPermissions, "Deselect all permissions (restrict)")
        Catch ex As Exception

        End Try

    End Sub

    'NDL: Added this function to enable cell-based tooltips for dgvGlobal
    'Modelled from https://msdn.microsoft.com/en-us/library/2249cf0a%28v=vs.85%29.aspx
    Sub dgvGlobal_CellMouseEnter(ByVal sender As Object, _
    ByVal e As DataGridViewCellEventArgs) _
    Handles dgvGlobal.CellMouseEnter
        Dim charConfigTitle
        Try
            If e.ColumnIndex = Me.dgvGlobal.Columns("CHARCONFIGTITLE").Index Then
                If ((e.ColumnIndex > -1) And (e.RowIndex > -1)) Then
                    charConfigTitle = Me.dgvGlobal.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                    With Me.dgvGlobal.Rows(e.RowIndex).Cells(e.ColumnIndex)
                        If (charConfigTitle.Equals("Enforce password integrity (TRUE or FALSE)")) Then
                            .ToolTipText = "Whether password must have 3 out of 4 of: (1) Upper case letter (2) Lower case letter (3) Digit (4) Non-alphanumeric character"
                        ElseIf (charConfigTitle.Equals("Minimum password length")) Then
                            .ToolTipText = "Minimum number of characters that must be in password"
                        ElseIf (charConfigTitle.Equals("Number of login attempts allowed")) Then
                            .ToolTipText = "Sequential login attempts allowed from within application"
                        ElseIf (charConfigTitle.Equals("Password change restriction (minutes)")) Then
                            .ToolTipText = ""  'Don't know what this does
                        ElseIf (charConfigTitle.Equals("Password expiration period (days)")) Then
                            .ToolTipText = "# of days users can use password before needing to change it"
                        ElseIf (charConfigTitle.Equals("Password history restriction")) Then
                            .ToolTipText = "For new password check: # of previous passwords to check for ""repeat"" passwords."
                        End If
                    End With
                End If
            End If
        Catch ex As Exception

        End Try
       
    End Sub


    Sub PermissionsItemChecked(lv As ListView, ByRef e As System.Windows.Forms.ItemCheckedEventArgs)


        If boolFormLoad Then
            Exit Sub
        End If

        If Me.dgvPermissions.RowCount = 0 Then
            Exit Sub
        End If

        Dim strLV As String

        Select Case lv.Name

            Case "lvPermissionsAdmin"
                strLV = "tblPermissionsAdmin"
            Case "lvPermissionsReportTemplate"
                strLV = "tblPermissionsReportTemplate"
            Case "lvPermissionsFinalReport"
                strLV = "tblPermissionsFinalReport"
            Case "lvPermissions"
                strLV = "tblPermissions"

        End Select

        Dim Count1 As Short
        Dim intRows As Short
        Dim boolC As Boolean
        Dim str1 As String
        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim strS As String
        Dim int1 As Short
        Dim lng1 As Int64

        Dim dtblP As System.Data.DataTable
        Dim rowsP() As DataRow
        Dim strF1 As String
        Dim strS1 As String
        Dim intRow As Short
        Dim dgv As DataGridView
        Dim strM As String

        dtblP = tblPermissions

        dgv = Me.dgvPermissions

        Dim dgvRow As DataGridViewRow

        dgvRow = dgv.SelectedRows(0)

        dtblP = tblPermissions
        lng1 = dgvRow.Cells("ID_TBLPERMISSIONS").Value

        If lng1 = 1 And boolPermLoad = False Then

            Dim boolF As Boolean = boolFormLoad

            If boolF Then
            Else
                strM = "The settings for the StudyDoc Administrator group may not be modified." & ChrW(10) & ChrW(10)
                strM = strM & "If it is desired to have an Administrator group to which changes can be made," & ChrW(10)
                strM = strM & "create a new Permissions Group and name it, for example, [CompanyName]_Administrator."
                MsgBox(strM, vbInformation, "Invalid action...")
            End If

            boolFormLoad = True
            e.Item.Checked = Not (e.Item.Checked)
            boolFormLoad = boolF

            'strM = "The settings for the StudyDoc Administrator group may not be modified." & ChrW(10) & ChrW(10)
            'strM = strM & "If it is desired to have an Administrator group to which changes can be made," & ChrW(10)
            'strM = strM & "create a new Permissions Group and name it, for example, [CompanyName]_Administrator."
            'MsgBox(strM, vbInformation, "Invalid action...")

            GoTo end1
        End If

        'If lng1 = 1 And e.Item.Checked = False And boolPermLoad = False Then
        '    e.Item.Checked = True
        '    strM = "The Administration settings for the Administrator group may not be modified."
        '    MsgBox(strM, vbInformation, "Invalid action...")
        '    GoTo end1
        'End If

        strF1 = "ID_TBLPERMISSIONS = " & lng1
        rowsP = dtblP.Select(strF1)
        If rowsP.Length = 0 Then
            Exit Sub
        End If

        dtbl = tblDataTableRowTitles
        'strF = "CHARDATATABLENAME = 'tblPermissionsAdmin' AND BOOLINCLUDE = -1 AND "
        strF = "CHARDATATABLENAME = '" & strLV & "' AND BOOLINCLUDE = -1 AND CHARTABLEREF IS NULL"
        strS = "INTORDER ASC"
        rows = dtbl.Select(strF, strS)
        int1 = rows.Length 'debugging

        Count1 = e.Item.Index

        str1 = rows(Count1).Item("CHARTABLEREFCOLUMNNAME")

        boolC = lv.Items(Count1).Checked
        rowsP(0).BeginEdit()
        If boolC Then
            'Me.dgvPermissions.Rows(0).Cells(str1).Value = -1
            rowsP(0).Item(str1) = -1
        Else
            'Me.dgvPermissions.Rows(0).Cells(str1).Value = 0
            rowsP(0).Item(str1) = 0
        End If
        rowsP(0).EndEdit()

        'Me.dgvPermissions.Rows(0).Cells(str1).Selected = True

end1:

    End Sub

    Private Sub lvPermissions_ItemChecked(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckedEventArgs) Handles lvPermissions.ItemChecked

        Call PermissionsItemChecked(Me.lvPermissions, e)

    End Sub


    Private Sub lvPermissionsAdmin_ItemChecked(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckedEventArgs) Handles lvPermissionsAdmin.ItemChecked

        Call PermissionsItemChecked(Me.lvPermissionsAdmin, e)

    End Sub

    Private Sub lvPermissionsReportTemplate_ItemChecked(sender As Object, e As ItemCheckedEventArgs) Handles lvPermissionsReportTemplate.ItemChecked

        Call PermissionsItemChecked(Me.lvPermissionsReportTemplate, e)

    End Sub

    Private Sub lvPermissionsFinalReport_ItemChecked(sender As Object, e As ItemCheckedEventArgs) Handles lvPermissionsFinalReport.ItemChecked

        Call PermissionsItemChecked(Me.lvPermissionsFinalReport, e)

    End Sub


    Private Sub dgvGlobal_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvGlobal.CellContentClick

    End Sub

    Function IsSelected(cbx As ComboBox) As Boolean

        IsSelected = True

        Dim intSI As Short 'selected index
        intSI = cbx.SelectedIndex
        If intSI = -1 Then 'ignore
            IsSelected = False
        End If


    End Function

    Private Sub cbxWatsonAccount_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxWatsonAccount.SelectedIndexChanged

        If IsSelected(Me.cbxWatsonAccount) Or boolFromWatsonButton Then
        Else
            Exit Sub
        End If

        If boolAccess Then
        Else

            Dim intID As Int64
            Dim dtbl As DataTable = tblWatsonUsers
            Dim strF As String
            Dim str1 As String = Me.cbxWatsonAccount.Text
            strF = "LOGINNAME = '" & str1 & "'"
            Dim rows() As DataRow = dtbl.Select(strF)
            Dim var1
            If rows.Length = 0 Then
                intID = 0
            Else
                intID = rows(0).Item("USERID")
            End If

            Me.ID_TBLWATSONACCOUNT.Text = intID

            If boolFormLoad Then
            Else
                'now update tblUserAccounts
                Dim idUA As Int64
                Dim dgv As DataGridView = Me.dgvUserAttributes
                Try
                    Dim intRow As Int16 = dgv.CurrentRow.Index
                    idUA = dgv("ID_TBLUSERACCOUNTS", intRow).Value
                    strF = "ID_TBLUSERACCOUNTS = " & idUA
                    Dim rowsUA() As DataRow = tblUserAccounts.Select(strF)
                    If rowsUA.Length = 0 Then
                    Else
                        rowsUA(0).BeginEdit()
                        rowsUA(0).Item("ID_TBLWATSONACCOUNT") = intID
                        rowsUA(0).EndEdit()
                    End If
                Catch ex As Exception
                    var1 = ex.Message
                End Try
            End If

        End If

end1:

    End Sub


    Private Sub cmdPopLDAP_Click(sender As Object, e As EventArgs) Handles cmdGetUserName.Click


        Call PopLDAP()

        'Select Case INTWINAUTH

        '    Case 1 'LDAP
        '        Call PopLDAP()
        '    Case 2 'Non-LDAP
        '        Dim boolA As Boolean = TestNetworkAccount()
        '    Case 3 'ADVAPI32
        '        Dim boolA As Boolean = TestNetworkAccount()
        'End Select

    End Sub


    Private Sub cmdTestAccount_Click(sender As Object, e As EventArgs) Handles cmdTestAccount.Click

        Dim boolA As Boolean = TestNetworkAccount()

    End Sub


    Function TestNetworkAccount() As Boolean

        TestNetworkAccount = False

        Dim strUN As String = Me.CHARNETWORKACCOUNT.Text
        Dim strLDAP As String = Me.CHARLDAP.Text
        Dim strPW As String
        Dim strM As String
        Dim boolA As Boolean

        If Len(strUN) = 0 Then
            strM = "Please enter network account userid in the cell."
            MsgBox(strM, vbInformation, "Invalid entry...")
            GoTo end2
        End If

        Select Case INTWINAUTH

            Case 1 'LDAP
                If Len(strLDAP) = 0 Then
                    strM = "Please enter an LDAP address in the LDAP cell."
                    MsgBox(strM, vbInformation, "Invalid entry...")
                    GoTo end2
                End If
  
        End Select

        Dim frm As New frmPasswordInput
        frm.ShowDialog()
        If frm.boolCancel Then
            GoTo end2
        End If

        strPW = frm.txtPassword.Text

        frm.Close()
        frm.Dispose()

        strM = "Enter Password:"

        'strPW = InputBox(strM, strM, "****")

        If Len(strPW) = 0 Then
            strM = "Password cannot be blank."
            MsgBox(strM, vbInformation, "Invalid entry...")
            GoTo end2
        End If

        Select Case INTWINAUTH

            Case 1 'LDAP
                TestNetworkAccount = AuthenticateUserLDAP(strLDAP, strUN, strPW)

            Case 2, 3 'Non-LDAP, ADVAPI32
                TestNetworkAccount = AuthenticateUser(strUN, strPW)

        End Select

end1:

        If TestNetworkAccount Then
            strM = "User Authenticated Successfully"
        Else
            strM = "User NOT Authenticated"
        End If
        MsgBox(strM, vbInformation, "Results...")

end2:



    End Function

    Sub PopLDAP()

        Dim frm As New frmLDAP
        Dim str1 As String
        Dim str2 As String
        Dim strM As String
        Dim intR As Short
        Dim strType As String = "Test"

        frm.txtLDAP.Text = Me.CHARLDAP.Text
        frm.strType = strType

        If Len(Me.LDAPUserID) = 0 Then
        Else
            frm.txtUserID.Text = Me.LDAPUserID
            frm.chkSaveCreds.Checked = True
        End If

        If Len(Me.LDAPPswd) = 0 Then
        Else
            frm.txtPswd.Text = Me.LDAPPswd
            frm.chkSaveCreds.Checked = True
        End If

        Call frm.FormLoad()

        frm.ShowDialog()

        If frm.chkSaveCreds.Checked Then
            Me.LDAPUserID = frm.txtUserID.Text
            Me.LDAPPswd = frm.txtPswd.Text
        End If

        Dim boolCancel As Boolean = frm.boolCancel
        If boolCancel Or frm.dgvUsers.RowCount = 0 Then
        Else

            Dim intRow As Int32 = frm.dgvUsers.CurrentRow.Index
            Dim strUID As String = frm.dgvUsers("samaccountname", intRow).Value
            Me.CHARNETWORKACCOUNT.Text = strUID
            Call SetCHARNETWORKACCOUNT()

            str1 = frm.txtLDAP.Text
            str2 = Me.CHARLDAP.Text

            If StrComp(str1, str2, CompareMethod.Text) = 0 Then
            Else
                strM = "The LDAP address in the 'LDAP Actions' form:" & ChrW(10) & ChrW(10) & "    " & frm.txtLDAP.Text
                strM = strM & ChrW(10) & ChrW(10) & "is different than the original LDAP address:" & ChrW(10) & ChrW(10) & "    " & Me.CHARLDAP.Text
                strM = strM & ChrW(10) & ChrW(10) & "Do wish to replace the original address with the address from the 'LDAP Actions' form?"
                intR = MsgBox(strM, vbYesNo, "Replace?")
                If intR = 6 Then
                    Me.CHARLDAP.Text = frm.txtLDAP.Text
                    Call SetCHARLDAP()
                End If
            End If

        End If

        frm.Dispose()

    End Sub

    Private Sub cmdClearWatson_Click(sender As Object, e As EventArgs) Handles cmdClearWatson.Click

        Dim intR As Short
        Dim strM As String
        Dim var1

        strM = "Do you wish to clear the assigned Watson account entry?"
        intR = MsgBox(strM, vbOKCancel, "Continue?")
        If intR = 1 Then
        Else
            GoTo end1
        End If

        Dim intID As Int64 = 0

        boolFromWatsonButton = True

        Me.ID_TBLWATSONACCOUNT.Text = intID
        Me.cbxWatsonAccount.SelectedIndex = -1

        boolFromWatsonButton = False

end1:

    End Sub

    Private Sub cmdClearLDAP_Click(sender As Object, e As EventArgs)

        Dim intR As Short
        Dim strM As String

        strM = "Do you wish to clear the assigned Windows network account entry?"
        intR = MsgBox(strM, vbOKCancel, "Continue?")
        If intR = 1 Then
        Else
            GoTo end1
        End If

        Me.CHARNETWORKACCOUNT.Clear()

end1:

    End Sub

    Private Sub CHARLDAP_Validated(sender As Object, e As EventArgs) Handles CHARLDAP.Validated

        Call SetCHARLDAP()

    End Sub

    Sub SetCHARLDAP()

        'Dim intID As Int64
        'Dim dtbl As DataTable = tblWatsonUsers
        'Dim strF As String
        'Dim str1 As String = Me.cbxWatsonAccount.Text
        'strF = "LOGINNAME = '" & str1 & "'"
        'Dim rows() As DataRow = dtbl.Select(strF)
        'Dim var1
        'If rows.Length = 0 Then
        '    intID = 0
        'Else
        '    intID = rows(0).Item("USERID")
        'End If

        'Me.ID_TBLWATSONACCOUNT.Text = intID

        Dim strF As String
        Dim var1

        var1 = Me.CHARLDAP.Text

        If boolFormLoad Then
        Else
            'now update tblUserAccounts
            Dim idUA As Int64
            Dim dgv As DataGridView = Me.dgvUserAttributes
            Try
                Dim intRow As Int16 = dgv.CurrentRow.Index
                idUA = dgv("ID_TBLUSERACCOUNTS", intRow).Value
                strF = "ID_TBLUSERACCOUNTS = " & idUA
                Dim rowsUA() As DataRow = tblUserAccounts.Select(strF)
                If rowsUA.Length = 0 Then
                Else
                    rowsUA(0).BeginEdit()
                    If Len(var1) = 0 Then
                        rowsUA(0).Item("CHARLDAP") = DBNull.Value
                    Else
                        rowsUA(0).Item("CHARLDAP") = var1
                    End If
                    rowsUA(0).EndEdit()
                End If
            Catch ex As Exception
                var1 = ex.Message
            End Try
        End If

    End Sub

    Private Sub CHARNETWORKACCOUNT_Validated(sender As Object, e As EventArgs) Handles CHARNETWORKACCOUNT.Validated

        Call SetCHARNETWORKACCOUNT()


    End Sub

    Sub SetCHARNETWORKACCOUNT()

        'Dim intID As Int64
        'Dim dtbl As DataTable = tblWatsonUsers
        'Dim strF As String
        'Dim str1 As String = Me.cbxWatsonAccount.Text
        'strF = "LOGINNAME = '" & str1 & "'"
        'Dim rows() As DataRow = dtbl.Select(strF)
        'Dim var1
        'If rows.Length = 0 Then
        '    intID = 0
        'Else
        '    intID = rows(0).Item("USERID")
        'End If

        'Me.ID_TBLWATSONACCOUNT.Text = intID

        Dim strF As String
        Dim var1

        var1 = Me.CHARNETWORKACCOUNT.Text

        If boolFormLoad Then
        Else
            'now update tblUserAccounts
            Dim idUA As Int64
            Dim dgv As DataGridView = Me.dgvUserAttributes
            Try
                Dim intRow As Int16 = dgv.CurrentRow.Index
                idUA = dgv("ID_TBLUSERACCOUNTS", intRow).Value
                strF = "ID_TBLUSERACCOUNTS = " & idUA
                Dim rowsUA() As DataRow = tblUserAccounts.Select(strF)
                If rowsUA.Length = 0 Then
                Else
                    rowsUA(0).BeginEdit()
                    If Len(var1) = 0 Then
                        rowsUA(0).Item("CHARNETWORKACCOUNT") = DBNull.Value
                    Else
                        rowsUA(0).Item("CHARNETWORKACCOUNT") = var1
                    End If
                    rowsUA(0).EndEdit()
                End If
            Catch ex As Exception
                var1 = ex.Message
            End Try
        End If


    End Sub

    Private Sub CHARLDAP_TextChanged(sender As Object, e As EventArgs) Handles CHARLDAP.TextChanged

    End Sub

    Private Sub cmdCopyLDAP_Click(sender As Object, e As EventArgs) Handles cmdCopyLDAP.Click


        Dim frm As New frmLDAP
        Dim str1 As String
        Dim str2 As String
        Dim strM As String
        Dim intR As Short
        Dim strType As String = "Existing"


        frm.strType = strType
        Call frm.FormLoad()

        'populate listbox with unique LDAPs
        Dim dv As DataView = New DataView(tblUserAccounts, "CHARLDAP IS NOT NULL", "CHARLDAP ASC", DataViewRowState.CurrentRows)
        Dim tbl1 As DataTable = dv.ToTable("a", True, "CHARLDAP")
        frm.lbxLDAP.DataSource = tbl1
        frm.lbxLDAP.DisplayMember = "CHARLDAP"
        'frm.lbxLDAP.DataBindings.Add(New Binding("USERID", dsDoPr, "TBLWATSONUSERS.USERID"))


        frm.ShowDialog()

        Dim boolCancel As Boolean = frm.boolCancel
        If boolCancel Or frm.lbxLDAP.Items.Count = 0 Then
        Else

            Try
                str1 = frm.lbxLDAP.Text
                Me.CHARLDAP.Text = str1
                Call SetCHARLDAP()
            Catch ex As Exception

            End Try
           

        End If

        frm.Dispose()

    End Sub

    Private Sub cmdClearNet_Click(sender As Object, e As EventArgs) Handles cmdClearNet.Click

        Dim intR As Short
        Dim strM As String

        strM = "Do you wish to clear the assigned network account entry?"
        intR = MsgBox(strM, vbOKCancel, "Continue?")
        If intR = 1 Then
        Else
            GoTo end1
        End If

        Me.CHARNETWORKACCOUNT.Clear()

        Call SetCHARNETWORKACCOUNT()

end1:

    End Sub

    Private Sub CHARNETWORKACCOUNT_TextChanged(sender As Object, e As EventArgs) Handles CHARNETWORKACCOUNT.TextChanged

    End Sub

    Private Sub rbLDAP_CheckedChanged(sender As Object, e As EventArgs) Handles rbLDAP.CheckedChanged

        Call LDAPDisplaySettings()

    End Sub

    Private Sub rbLDAPNon_CheckedChanged(sender As Object, e As EventArgs) Handles rbLDAPNon.CheckedChanged

        Call LDAPDisplaySettings()

    End Sub

    Private Sub rbADVAPI32_CheckedChanged(sender As Object, e As EventArgs) Handles rbADVAPI32.CheckedChanged

        Call LDAPDisplaySettings()

    End Sub

    Sub LDAPDisplaySettings()

        If boolFormLoad Then
            Exit Sub
        End If

        Dim boolVis As Boolean = True

        If Me.rbLDAP.Checked Then

            INTWINAUTH = 1
            boolVis = True
            Me.cmdGetUserName.Visible = True

        ElseIf Me.rbLDAPNon.Checked Then

            INTWINAUTH = 2
            boolVis = False
            Me.cmdGetUserName.Visible = False

        ElseIf Me.rbADVAPI32.Checked Then

            INTWINAUTH = 3
            boolVis = False
            Me.cmdGetUserName.Visible = False

        End If

        Me.panLDAP.Visible = boolVis

    End Sub

    Sub SaveLDAP()

        Dim strF As String = "CHARCONFIGCATEGORY = 'Global LDAP'"
        Dim rows() As DataRow = tblConfiguration.Select(strF)
        Dim var1

        If rows.Length = 0 Then
        Else

            rows(0).BeginEdit()
            rows(0).Item("CHARCONFIGVALUE") = INTWINAUTH
            rows(0).EndEdit()

        End If

    End Sub

    Sub SetLDAP()

        Dim strF As String = "CHARCONFIGCATEGORY = 'Global LDAP'"
        Dim rows() As DataRow = tblConfiguration.Select(strF)
        Dim var1

        If rows.Length = 0 Then
            INTWINAUTH = 3
        Else
            var1 = NZ(rows(0).Item("CHARCONFIGVALUE"), 1)
            If IsNumeric(var1) Then
                INTWINAUTH = var1
            End If
        End If

        Select Case INTWINAUTH
            Case 1
                Me.rbLDAP.Checked = True
            Case 2
                Me.rbLDAPNon.Checked = True
            Case 3
                Me.rbADVAPI32.Checked = True
        End Select

    End Sub


End Class