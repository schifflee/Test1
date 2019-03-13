Option Compare Text

Public Class frmAddSD

    Public strMod As String = ""
    Public intC As Integer = 1
    Public arrC(10, 100)
    '1=ColumnName,2=HeaderText,3=datatype,4=value for id's, 5=id for id's
    Public tblT As System.Data.DataTable
    Public tbl1 As New System.Data.DataTable
    Public tbl2 As New System.Data.DataTable
    Public tbl3 As New System.Data.DataTable
    Public tbl4 As New System.Data.DataTable
    Public tbl5 As New System.Data.DataTable
    Public boolCancel As Boolean = True
    Public boolFormLoad As Boolean = False
    Public maxIDProj As Int64 = 0
    Public maxIDStudy As Int64 = 0

    Sub Filldgv1(ByVal dgv2 As DataGridView)

        Dim dgv1 As DataGridView
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
        Dim dtbl As New System.Data.DataTable
        Dim bool As Boolean
        Dim strID As String
        Dim intRow As Short
        Dim boolDo As Boolean

        dgv1 = Me.dgv1
        intCols = dgv2.Columns.Count
        intRows = dgv2.Rows.Count
        intRow = frmSD.dgvSDProjectS.CurrentRow.Index

        int1 = 0
        Select Case strMod
            Case "Projects"

                For Count1 = 0 To intRows - 1
                    If dgv2.Rows(Count1).Cells("ColumnValue").ReadOnly Then
                    Else
                        int1 = int1 + 1
                        var1 = dgv2.Rows(Count1).Cells("ColumnName").Value
                        arrC(1, int1) = dgv2.Rows(Count1).Cells("ColumnName").Value
                        var1 = dgv2.Rows(Count1).Cells("ColumnHeader").Value
                        arrC(2, int1) = dgv2.Rows(Count1).Cells("ColumnHeader").Value
                        var1 = dgv2.Rows(Count1).Cells("ColumnValue").Value
                        arrC(4, int1) = dgv2.Rows(Count1).Cells("ColumnValue").Value
                    End If
                Next

            Case "Studies"

                For Count1 = 0 To intRows - 1
                    boolDo = False
                    str1 = dgv2.Rows(Count1).Cells("ColumnName").Value
                    If StrComp(str1, "ID_TBLGUWUSTUDIES", CompareMethod.Text) = 0 Or IsUserIDStuff(str1) Then
                    Else
                        strID = Mid(str1, 1, 3)
                        var1 = dgv2.Rows(Count1).Cells("ColumnReadOnly").Value
                        bool = CBool(dgv2.Rows(Count1).Cells("ColumnReadOnly").Value)
                        'If dgv2.Rows(Count1).Cells("ColumnName").ReadOnly = False Or StrComp(strID, "ID_", CompareMethod.Text) = 0 Then
                        If bool = False Or StrComp(strID, "ID_", CompareMethod.Text) = 0 Then
                            'bool = dgv2.Rows(Count1).Cells("ColumnName").ReadOnly

                            If StrComp(str1, "ID_TBLGUWUPROJECTS", CompareMethod.Text) = 0 Then 'continue
                                boolDo = True
                            Else

                                If bool And StrComp(str1, "ID_TBLGUWUSTUDIES", CompareMethod.Text) <> 0 Then
                                    boolDo = False
                                Else
                                    boolDo = True
                                End If

                            End If
                            If boolDo Then

                                int1 = int1 + 1
                                var1 = dgv2.Rows(Count1).Cells("ColumnName").Value
                                arrC(1, int1) = dgv2.Rows(Count1).Cells("ColumnName").Value
                                var1 = dgv2.Rows(Count1).Cells("ColumnHeader").Value
                                arrC(2, int1) = dgv2.Rows(Count1).Cells("ColumnHeader").Value
                                var1 = dgv2.Rows(Count1).Cells("ColumnValue").Value
                                arrC(4, int1) = dgv2.Rows(Count1).Cells("ColumnValue").Value
                                If StrComp(str1, "ID_TBLGUWUPROJECTS", CompareMethod.Text) = 0 Then
                                    'enter Project Number and ID
                                    var1 = frmSD.dgvSDProjectS.Item("CHARPROJECTNUM", intRow).Value
                                    var2 = frmSD.dgvSDProjectS.Item("ID_TBLGUWUPROJECTS", intRow).Value
                                    arrC(4, int1) = var1
                                    arrC(5, int1) = var2
                                ElseIf StrComp(str1, "BOOLISGLP", CompareMethod.Text) = 0 Then
                                    arrC(4, int1) = "FALSE"
                                    arrC(5, int1) = 0
                                End If
                            End If
                        End If
                    End If

                    'If dgv2.Rows(Count1).Cells("ColumnValue").ReadOnly Then
                    'Else
                    '    int1 = int1 + 1
                    '    var1 = dgv2.Rows(Count1).Cells("ColumnName").Value
                    '    arrC(1, int1) = dgv2.Rows(Count1).Cells("ColumnName").Value
                    '    var1 = dgv2.Rows(Count1).Cells("ColumnHeader").Value
                    '    arrC(2, int1) = dgv2.Rows(Count1).Cells("ColumnHeader").Value
                    '    var1 = dgv2.Rows(Count1).Cells("ColumnValue").Value
                    '    arrC(4, int1) = dgv2.Rows(Count1).Cells("ColumnValue").Value
                    'End If
                Next

                'For Count1 = 0 To intCols - 1
                '    str1 = dgv2.Columns(Count1).Name
                '    If StrComp(str1, "ID_TBLGUWUSTUDIES", CompareMethod.Text) = 0 Then
                '    Else
                '        strID = Mid(str1, 1, 3)
                '        If dgv2.Columns(Count1).Visible Or StrComp(strID, "ID_", CompareMethod.Text) = 0 Then
                '            bool = dgv2.Columns(Count1).ReadOnly
                '            If bool And StrComp(str1, "ID_TBLGUWUSTUDIES", CompareMethod.Text) <> 0 Then
                '            Else
                '                int1 = int1 + 1
                '                arrC(1, int1) = dgv2.Columns(Count1).Name
                '                arrC(2, int1) = dgv2.Columns(Count1).HeaderText
                '                If StrComp(str1, "ID_TBLGUWUPROJECTS", CompareMethod.Text) = 0 Then
                '                    'enter Project Number and ID
                '                    var1 = frmSD.dgvSDProjectS.Item("CHARPROJECTNUM", intRow).Value
                '                    var2 = frmSD.dgvSDProjectS.Item("ID_TBLGUWUPROJECTS", intRow).Value
                '                    arrC(4, int1) = var1
                '                    arrC(5, int1) = var2
                '                End If
                '            End If
                '        End If
                '    End If
                'Next

        End Select
        intRows = int1

        For Count1 = 1 To intRows
            str1 = arrC(1, Count1)
            var1 = tblT.Columns(str1).DataType.ToString
            arrC(3, Count1) = var1
            '''''''''console.writeline(var1.ToString)
        Next

        'make new datatable
        Select Case intC
            Case 1 'Projects
                dtbl = tbl1
            Case 2 'Studies
                dtbl = tbl2
            Case 3
                dtbl = tbl3
            Case 4
                dtbl = tbl4
            Case 5
                dtbl = tbl5
        End Select

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

        For Count1 = 1 To intRows
            Dim row As DataRow = dtbl.NewRow
            row("ColumnName") = arrC(1, Count1)
            row("ColumnHeader") = arrC(2, Count1)
            'row("ColumnValue") = arrC(4, Count1)
            dtbl.Rows.Add(row)
        Next

        Dim dv As System.Data.DataView = New DataView(dtbl)
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

                'ElseIf StrComp(str4, "BOOL", CompareMethod.Text) = 0 Then
                '    Select Case Me.strMod
                '        Case "Studies"
                '            dgv1("ColumnValue", Count1 - 1).Value = "FALSE"
                '    End Select
            ElseIf StrComp(str1, "BOOLISGLP", CompareMethod.Text) = 0 And StrComp(strMod, "Studies", CompareMethod.Text) = 0 Then
                dgv1("ColumnValue", Count1 - 1).Value = arrC(4, Count1)
                dgv1("ColumnID", Count1 - 1).Value = arrC(5, Count1)
            ElseIf StrComp(str1, "ID_TBLGUWUPROJECTS", CompareMethod.Text) = 0 And StrComp(strMod, "Studies", CompareMethod.Text) = 0 Then
                dgv1("ColumnValue", Count1 - 1).Value = arrC(4, Count1)
                dgv1("ColumnID", Count1 - 1).Value = arrC(5, Count1)
            End If

        Next

        dgv1.Columns("ColumnName").Visible = False
        dgv1.Columns("ColumnHeader").HeaderText = "Item"
        dgv1.Columns("ColumnHeader").ReadOnly = True
        dgv1.Columns("ColumnValue").HeaderText = "Value"
        dgv1.Columns("ColumnValue").Visible = True
        dgv1.Columns("ColumnID").Visible = False 'False

        dgv1.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
        dgv1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        dgv1.AllowUserToResizeColumns = True
        dgv1.AllowUserToResizeRows = True
        dgv1.RowHeadersWidth = 25
        dgv1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        dgv1.DefaultCellStyle.WrapMode = DataGridViewTriState.True

        dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        dgv1.Columns.Item("ColumnValue").MinimumWidth = 200
        Try
            dgv1.Columns.Item("ColumnHeader").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            dgv1.Columns.Item("ColummValue").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        Catch ex As Exception

        End Try

        dgv1.AutoResizeColumns()

        If dgv1.Rows.Count = 0 Then
        Else
            dgv1.CurrentCell = dgv1.Item("ColumnValue", 0)
        End If


end1:

    End Sub

    Sub Filldgv1BU(ByVal dgv2 As DataGridView)

        Dim dgv1 As DataGridView
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
        Dim dtbl As New System.Data.DataTable
        Dim bool As Boolean
        Dim strID As String
        Dim intRow As Short

        dgv1 = Me.dgv1
        intCols = dgv2.Columns.Count
        intRow = frmSD.dgvSDProjectS.CurrentRow.Index

        int1 = 0
        Select Case strMod
            Case "Projects"
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

            Case "Studies"
                For Count1 = 0 To intCols - 1
                    str1 = dgv2.Columns(Count1).Name
                    If StrComp(str1, "ID_TBLGUWUSTUDIES", CompareMethod.Text) = 0 Then
                    Else
                        strID = Mid(str1, 1, 3)
                        If dgv2.Columns(Count1).Visible Or StrComp(strID, "ID_", CompareMethod.Text) = 0 Then
                            bool = dgv2.Columns(Count1).ReadOnly
                            If bool And StrComp(str1, "ID_TBLGUWUSTUDIES", CompareMethod.Text) <> 0 Then
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

        End Select
        intRows = int1

        For Count1 = 1 To intRows
            str1 = arrC(1, Count1)
            var1 = tblT.Columns(str1).DataType.ToString
            arrC(3, Count1) = var1
            '''''''''console.writeline(var1.ToString)
        Next

        'make new datatable
        Select Case intC
            Case 1 'Projects
                dtbl = tbl1
            Case 2 'Studies
                dtbl = tbl2
            Case 3
                dtbl = tbl3
            Case 4
                dtbl = tbl4
            Case 5
                dtbl = tbl5
        End Select

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

        For Count1 = 1 To intRows
            Dim row As DataRow = dtbl.NewRow
            row("ColumnName") = arrC(1, Count1)
            row("ColumnHeader") = arrC(2, Count1)
            dtbl.Rows.Add(row)
        Next

        Dim dv As System.Data.DataView = New DataView(dtbl)
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
                Select Case Me.strMod
                    Case "Studies"
                        dgv1("ColumnValue", Count1 - 1).Value = "FALSE"
                End Select
            ElseIf StrComp(str1, "ID_TBLGUWUPROJECTS", CompareMethod.Text) = 0 And StrComp(strMod, "Studies", CompareMethod.Text) = 0 Then
                dgv1("ColumnValue", Count1 - 1).Value = arrC(4, Count1)
                dgv1("ColumnID", Count1 - 1).Value = arrC(5, Count1)
            End If

        Next

        dgv1.Columns("ColumnName").Visible = False
        dgv1.Columns("ColumnHeader").HeaderText = "Item"
        dgv1.Columns("ColumnHeader").ReadOnly = True
        dgv1.Columns("ColumnValue").HeaderText = "Value"
        dgv1.Columns("ColumnID").Visible = False

        dgv1.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
        dgv1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        dgv1.AllowUserToResizeColumns = True
        dgv1.AllowUserToResizeRows = True
        dgv1.RowHeadersWidth = 25
        dgv1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        dgv1.DefaultCellStyle.WrapMode = DataGridViewTriState.True

        dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        dgv1.Columns.Item("ColumnValue").MinimumWidth = 200
        Try
            dgv1.Columns.Item("ColumnHeader").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            dgv1.Columns.Item("ColummValue").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        Catch ex As Exception

        End Try

        dgv1.AutoResizeColumns()

        If dgv1.Rows.Count = 0 Then
        Else
            dgv1.CurrentCell = dgv1.Item("ColumnValue", 0)
        End If


end1:

    End Sub

    Sub ForceCellFormat()
        Dim str1 As String
        Dim str2 As String
        Dim intRow As Short
        Dim dgv As DataGridView
        Dim strBool As String
        Dim strDt As String
        Dim locX, locY
        Dim var1
        Dim rows() As DataRow
        Dim strF As String
        Dim strS As String

        dgv = Me.dgv1
        Me.mCal1.Visible = False

        intRow = dgv.CurrentRow.Index
        str1 = dgv.Rows(intRow).Cells("ColumnName").Value
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
            cbx.Value = "FALSE"
            cbx.DisplayStyleForCurrentCellOnly = True
            cbx.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
            dgv("ColumnValue", intRow) = cbx

        ElseIf StrComp(strDt, "DT", CompareMethod.Text) = 0 Then
            locX = Cursor.Position.X
            locY = Cursor.Position.Y
            locX = dgv.Left + dgv.RowHeadersWidth + (dgv.Columns(1).Width * 2)
            locY = dgv.Location.Y + (dgv.Rows(intRow).Height * intRow) + dgv.ColumnHeadersHeight
            If IsDate(var1) Then
                Me.mCal1.SelectionStart = var1
                Me.mCal1.SelectionEnd = var1
            Else
                Me.mCal1.SelectionStart = Now
                Me.mCal1.SelectionEnd = Now
            End If
            Me.mCal1.Location = New System.Drawing.Point(locX, locY)
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
            strF = "ID_tblGuWuStudyStat > -1"
            strS = "ID_TBLGUWUSTUDYSTAT ASC"
            rows = tblGuWuStudyStat.Select(strF, strS)
            cbx.DataSource = rows
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
            strF = "ID_TBLGUWUSTUDYDESIGNTYPE > -1"
            strS = "ID_TBLGUWUSTUDYDESIGNTYPE ASC"
            rows = tblGuWuStudyDesignType.Select(strF, strS)
            cbx.DataSource = rows 'tblGuWuStudyDesignType
            cbx.DisplayMember = tblGuWuStudyDesignType.Columns("CHARSTUDYDESIGNTYPE").ColumnName
            If Len(var1) = 0 Then
            Else
                cbx.Value = var1
            End If
            cbx.DisplayStyleForCurrentCellOnly = True
            cbx.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton

            dgv("ColumnValue", intRow) = cbx

        End If

    End Sub

    Function AcceptProject() As Boolean

        Dim dtbl As System.Data.DataTable
        Dim dgv As DataGridView
        Dim dv As System.Data.DataView
        Dim intRows As Short
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim var1
        Dim boolGo As Boolean
        Dim strM As String
        Dim strF As String
        Dim maxID As Int64

        AcceptProject = False

        dgv = Me.dgv1
        intRows = dgv.Rows.Count

        'first validate
        boolGo = True
        strM = ""
        For Count1 = 0 To intRows - 1
            str1 = dgv.Item("ColumnName", Count1).Value
            str2 = dgv.Item("ColumnHeader", Count1).Value
            var1 = dgv.Item("ColumnValue", Count1).Value
            Select Case str1
                Case "CHARPROJECTNAME"
                    If Len(NZ(var1, "")) = 0 Then
                        boolGo = False
                        Exit For
                    End If
                Case "CHARPROJECTNUM"
                    If Len(NZ(var1, "")) = 0 Then
                        boolGo = False
                        Exit For
                    End If
                Case "CHARPROJECTDESCR"

            End Select
        Next

        If boolGo Then
        Else
            strM = "Row '" & str2 & "' must contain a value."
            dgv.CurrentCell = dgv.Rows(Count1).Cells("ColumnValue")
            MsgBox(strM, MsgBoxStyle.Information, "Invalid value...")
            GoTo end1
        End If

        dtbl = tblGuWuProjects

        Dim row As DataRow = dtbl.NewRow
        maxID = GetMaxID("TBLGUWUPROJECTS", 1, True)
        maxIDProj = maxID
        frmSD.SDProjAddID = maxIDProj
        row.BeginEdit()
        row.Item("ID_TBLGUWUPROJECTS") = maxID
        For Count1 = 0 To intRows - 1
            str1 = dgv.Item("ColumnName", Count1).Value
            var1 = dgv.Item("ColumnValue", Count1).Value
            row.Item(str1) = var1
        Next
        row.EndEdit()
        dtbl.Rows.Add(row)

        'don't do this anymore.
        'change filter in home dgv
        'dgv = frmSD.dgvSDProject
        'strF = "ID_TBLGUWUPROJECTS = " & maxID
        'dv = dgv.DataSource
        'dv.RowFilter = strF

        'dgv = frmSD.dgvSDProjectS
        'strF = "ID_TBLGUWUPROJECTS = " & maxID
        'dv = dgv.DataSource
        'dv.RowFilter = strF

        'select the row instead
        Dim int1 As Int64
        Dim id As Int64
        Dim intRow As Int64
        dgv = frmSD.dgvSDProject
        dv = dgv.DataSource
        int1 = dgv.Rows.Count
        intRow = -1
        For Count1 = 0 To int1 - 1
            id = dv(Count1).Item("ID_TBLGUWUPROJECTS")
            If id = maxID Then
                intRow = Count1
                Exit For
            End If
        Next
        dgv.CurrentCell = dgv.Rows(intRow).Cells("CHARPROJECTNUM")
        dgv.CurrentRow.Selected = True

        dgv = frmSD.dgvSDProjectS
        dv = dgv.DataSource
        int1 = dgv.Rows.Count
        intRow = -1
        For Count1 = 0 To int1 - 1
            id = dv(Count1).Item("ID_TBLGUWUPROJECTs")
            If id = maxID Then
                intRow = Count1
                Exit For
            End If
        Next
        dgv.CurrentCell = dgv.Rows(intRow).Cells("CHARPROJECTNUM")
        dgv.CurrentRow.Selected = True


        AcceptProject = True

end1:

    End Function

    Function AcceptStudy() As Boolean

        Dim dtbl As System.Data.DataTable
        Dim dgv As DataGridView
        Dim dv As System.Data.DataView
        Dim intRows As Short
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim var1, var2
        Dim boolGo As Boolean
        Dim strM As String
        Dim strF As String
        Dim boolDt As Boolean
        Dim maxID As Int64
        Dim strBool As String

        AcceptStudy = False

        dgv = Me.dgv1
        intRows = dgv.Rows.Count

        'first validate
        boolGo = True
        strM = ""
        For Count1 = 0 To intRows - 1
            str1 = dgv.Item("ColumnName", Count1).Value
            str2 = dgv.Item("ColumnHeader", Count1).Value
            var1 = dgv.Item("ColumnValue", Count1).Value
            var2 = dgv.Item("ColumnID", Count1).Value
            boolDt = False
            Select Case str1

                Case "ID_TBLGUWUPROJECTS"


                Case "CHARSTUDYNAME"
                    If Len(NZ(var1, "")) = 0 Then
                        boolGo = False
                        Exit For
                    End If
                Case "CHARSTUDYNUMBER"
                    If Len(NZ(var1, "")) = 0 Then
                        boolGo = False
                        Exit For
                    End If

                Case "CHARSTUDYDESCR"
                    If Len(NZ(var1, "")) = 0 Then
                        boolGo = True
                        Exit For
                    End If

                Case "BOOLISGLP"
                    If Len(NZ(var1, "")) = 0 Then
                        boolGo = False
                        Exit For
                    End If

                Case "DTSTUDYSTARTPRE"
                    boolDt = True
                    If Len(NZ(var1, "")) = 0 Then
                    ElseIf IsDate(var1) Then
                    Else
                        boolGo = False
                        Exit For
                    End If

                Case "DTSTUDYSTARTACT"
                    boolDt = True
                    If Len(NZ(var1, "")) = 0 Then
                    ElseIf IsDate(var1) Then
                    Else
                        boolGo = False
                        Exit For
                    End If


                Case "DTSTUDYENDPRED"
                    boolDt = True
                    If Len(NZ(var1, "")) = 0 Then
                    ElseIf IsDate(var1) Then
                    Else
                        boolGo = False
                        Exit For
                    End If


                Case "DTSTUDYENDACT"
                    boolDt = True
                    If Len(NZ(var1, "")) = 0 Then
                    ElseIf IsDate(var1) Then
                    Else
                        boolGo = False
                        Exit For
                    End If


                Case "DTEXTRACTIONDATE"
                    boolDt = True
                    If Len(NZ(var1, "")) = 0 Then
                    ElseIf IsDate(var1) Then
                    Else
                        boolGo = False
                        Exit For
                    End If


                Case "CHARNOTEBOOKREF"


            End Select
        Next

        If boolGo Then
        Else
            If boolDt Then
                strM = "Row '" & str2 & "' must contain a date value."
            Else
                strM = "Row '" & str2 & "' must contain a value."
            End If
            dgv.CurrentCell = dgv.Rows(Count1).Cells("ColumnValue")
            MsgBox(strM, MsgBoxStyle.Information, "Invalid value...")
            GoTo end1
        End If


        dtbl = tblGuWuStudies

        Dim strID As String
        Dim row As DataRow = dtbl.NewRow
        maxID = GetMaxID("TBLGUWUSTUDIES", 1, True)
        maxIDStudy = maxID
        frmSD.SDStudyAddID = maxIDStudy
        row.BeginEdit()
        row.Item("ID_TBLGUWUSTUDIES") = maxID
        For Count1 = 0 To intRows - 1
            str1 = dgv.Item("ColumnName", Count1).Value
            strID = Mid(str1, 1, 3)
            strBool = Mid(str1, 1, 4)
            If StrComp(str1, "ID_TBLGUWUPROJECTS", CompareMethod.Text) = 0 Then
                var1 = dgv.Item("ColumnID", Count1).Value
            ElseIf StrComp(strID, "ID_", CompareMethod.Text) = 0 Then
                var1 = dgv.Item("ColumnID", Count1).Value
            ElseIf StrComp(strBool, "BOOL", CompareMethod.Text) = 0 Then
                var1 = dgv.Item("ColumnID", Count1).Value
            Else
                var1 = dgv.Item("ColumnValue", Count1).Value
            End If
            row.Item(str1) = var1
        Next
        row.EndEdit()
        dtbl.Rows.Add(row)


        'change SELECTION in home dgv
        Dim int1 As Int64
        Dim id As Int64
        Dim intRow As Int64

        dgv = frmSD.dgvSDStudy
        'strF = "ID_TBLGUWUSTUDIES = " & maxID
        'dv = dgv.DataSource
        'dv.RowFilter = strF

        'select the row instead
        dv = dgv.DataSource
        int1 = dgv.Rows.Count
        intRow = -1
        For Count1 = 0 To int1 - 1
            id = dv(Count1).Item("ID_TBLGUWUSTUDIES")
            If id = maxID Then
                intRow = Count1
                Exit For
            End If
        Next
        dgv.CurrentCell = dgv.Rows(intRow).Cells("CHARSTUDYNUMBER")
        dgv.CurrentRow.Selected = True

        AcceptStudy = True

end1:

    End Function

    Private Sub frmAddSD_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub cmdOK1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK1.Click
        Select Case strMod
            Case "Projects"
                If AcceptProject() Then
                    frmSD.cmdAddProject.Enabled = False
                    frmSD.boolSDProjAdd = True
                    Me.boolCancel = False
                    Me.Visible = False
                End If

            Case "Studies"
                If AcceptStudy() Then
                    frmSD.cmdAddStudy.Enabled = False
                    frmSD.boolSDStudyAdd = True
                    Me.boolCancel = False
                    Me.Visible = False
                End If

        End Select


    End Sub

    Private Sub cmdCancel1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCancel1.Click
        Me.boolCancel = True
        Me.Visible = False
    End Sub

    Private Sub dgv1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv1.CellClick

        If e.ColumnIndex = 2 Then
        Else
            GoTo end1
        End If

        Call ForceCellFormat()

end1:

    End Sub

    Private Sub mCal1_DateSelected(ByVal sender As Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles mCal1.DateSelected

        Dim dgv As DataGridView
        Dim intRow As Short

        dgv = Me.dgv1
        intRow = dgv.CurrentRow.Index

        dgv.Rows(intRow).Cells("ColumnValue").Value = e.Start

    End Sub

    Private Sub dgv1_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgv1.CellValidating

        If boolFormLoad Then
            Exit Sub
        End If

        Dim dgv As DataGridView
        Dim intRow As Short
        Dim var1
        Dim strF As String
        Dim rows() As DataRow
        Dim dtbl As System.Data.DataTable
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim boolGo As Boolean
        Dim boolHit As Boolean
        Dim strM As String

        dgv = Me.dgv1
        intRow = dgv.CurrentRow.Index
        str1 = dgv("ColumnName", intRow).Value
        str3 = dgv("ColumnHeader", intRow).Value
        'var1 = NZ(dgv("ColumnValue", intRow).Value, "")
        var1 = NZ(e.FormattedValue, "")

        If e.ColumnIndex <> 2 Then
            Exit Sub
        End If

        boolGo = False
        If StrComp(strMod, "Projects", CompareMethod.Text) = 0 Then
            If StrComp(str1, "CHARPROJECTNAME", CompareMethod.Text) = 0 Then
                boolGo = True
                strF = "CHARPROJECTNAME = '" & NZ(var1, "aaGubbsJunckCrapaa") & "'"
            ElseIf StrComp(str1, "CHARPROJECTNUM", CompareMethod.Text) = 0 Then
                boolGo = True
                strF = "CHARPROJECTNUM = '" & NZ(var1, "aaGubbsJunckCrapaa") & "'"
            End If
            If boolGo Then
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
                strF = "CHARSTUDYNAME = '" & NZ(var1, "aaGubbsJunckCrapaa") & "'"
            ElseIf StrComp(str1, "CHARSTUDYNUM", CompareMethod.Text) = 0 Then
                boolGo = True
                strF = "CHARSTUDYNUM = '" & NZ(var1, "aaGubbsJunckCrapaa") & "'"
            End If
            If boolGo Then
                dtbl = tblGuWuStudies
                rows = dtbl.Select(strF)
                If rows.Length = 0 Then 'OK
                Else
                    boolHit = True
                End If
            End If

            Call MoreCellValidating(e.FormattedValue, "Studies")

        End If


        If boolHit Then
            e.Cancel = True
            dgv.CurrentCell = dgv.Rows(intRow).Cells("ColumnValue")
            strM = "The field '" & str3 & "' does not allow duplicates." & ChrW(10) & "Please enter a different value."
            MsgBox(strM, MsgBoxStyle.Information, "Duplicates not allowed...")
        End If

    End Sub

    Sub MoreCellValidating(ByVal varValue As Object, ByVal strMod As String)

        'varValue is e.formattedvalue

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

        dgv = Me.dgv1
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
            If StrComp(varValue, "TRUE", CompareMethod.Text) = 0 Then
                dgv.Item("ColumnID", intRow).Value = -1
            Else
                dgv.Item("ColumnID", intRow).Value = 0
            End If
        ElseIf boolID Then
            'If Len(varValue) = 0 Then
            'Else
            If StrComp(str1, "ID_TBLGUWUPROJECTS", CompareMethod.Text) = 0 Then


            ElseIf StrComp(str1, "ID_TBLCONFIGREPORTTYPE", CompareMethod.Text) = 0 Then
                strF = "CHARREPORTTYPE = '" & varValue & "'"
                rows = tblConfigReportType.Select(strF)
                int1 = rows(0).Item("ID_TBLCONFIGREPORTTYPE")
                If int1 = 0 Then
                    dgv.Item("ColumnID", intRow).Value = DBNull.Value
                Else
                    dgv.Item("ColumnID", intRow).Value = rows(0).Item("ID_TBLCONFIGREPORTTYPE")
                End If

            ElseIf StrComp(str1, "ID_TBLGUWUSTUDYSTAT", CompareMethod.Text) = 0 Then
                strF = "CHARSTATUS = '" & varValue & "'"
                rows = tblGuWuStudyStat.Select(strF)
                int1 = rows(0).Item("ID_TBLGUWUSTUDYSTAT")
                If int1 = 0 Then
                    dgv.Item("ColumnID", intRow).Value = DBNull.Value
                Else
                    dgv.Item("ColumnID", intRow).Value = rows(0).Item("ID_TBLGUWUSTUDYSTAT")
                End If

            ElseIf StrComp(str1, "ID_TBLGUWUSTUDYDESIGNTYPE", CompareMethod.Text) = 0 Then
                strF = "CHARSTUDYDESIGNTYPE = '" & varValue & "'"
                rows = tblGuWuStudyDesignType.Select(strF)
                int1 = rows(0).Item("ID_TBLGUWUSTUDYDESIGNTYPE")
                If int1 = 0 Then
                    dgv.Item("ColumnID", intRow).Value = DBNull.Value
                Else
                    dgv.Item("ColumnID", intRow).Value = rows(0).Item("ID_TBLGUWUSTUDYDESIGNTYPE")
                End If

            End If
        End If

        'End If

end1:
    End Sub


End Class