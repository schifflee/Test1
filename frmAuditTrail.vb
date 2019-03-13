Option Compare Text

Imports Microsoft.VisualBasic.PowerPacks
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Imports System
Imports System.IO
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Documents
Imports System.Windows.Media

Public Class frmAuditTrail

    Private checkPrint As Integer
    Private boolHold As Boolean = False
    Private boolNeedUpdate As Boolean = False
    Public strForm As String = ""
    Private boolFormLoad As Boolean = False

    Private Sub PrintDocument1_BeginPrint(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintEventArgs) Handles PrintDocument1.BeginPrint
        checkPrint = 0
    End Sub


    Sub FillSort()

        Dim cbx As Windows.Forms.ComboBox
        Dim Count1 As Short

        For Count1 = 1 To 4
            Select Case Count1
                Case 1
                    cbx = Me.cbxSort1
                Case 2
                    cbx = Me.cbxSort2
                Case 3
                    cbx = Me.cbxSort3
                Case 4
                    cbx = Me.cbxSort4

            End Select

            cbx.Items.Add("")
            cbx.Items.Add("Study Name")
            cbx.Items.Add("User Name")
            cbx.Items.Add("Changed Item")
            cbx.Items.Add("Event Date")

            If Count1 = 1 Then
                cbx.SelectedIndex = 4
            Else
                cbx.SelectedIndex = -1
            End If

        Next

        Me.cbxFilterSource.Items.Add("Study Name")
        Me.cbxFilterSource.Items.Add("User Name")
        Me.cbxFilterSource.Items.Add("Report Writer Administration")
        Me.cbxFilterSource.Items.Add("StudyDoc Administration")
        Me.cbxFilterSource.Items.Add("Microsoft" & ChrW(8482) & " Word Templates")

        Select Case strForm
            Case "frmConsole"
                Me.cbxFilterSource.SelectedIndex = 3
            Case "frmHome_01"
                Me.cbxFilterSource.SelectedIndex = 0
            Case Else
                Me.cbxFilterSource.SelectedIndex = 0
        End Select


    End Sub

    Sub DoDateFilters()

        If Me.chkFilterDates.Checked Then
            Me.DateTimePicker1.Enabled = True
            Me.DateTimePicker2.Enabled = True
        Else
            Me.DateTimePicker1.Enabled = False
            Me.DateTimePicker2.Enabled = False
        End If


    End Sub

    Sub OpenConnection()

        Cursor.Current = Cursors.WaitCursor

        Dim dt1 As Date
        Dim dt2 As Date
        Dim str1 As String

        dt1 = Me.DateTimePicker1.Value
        dt2 = Me.DateTimePicker2.Value

        dt1 = DateAdd(DateInterval.Day, -1, dt1)
        dt2 = DateAdd(DateInterval.Day, 1, dt2)

        Try

            '20160531 LEE: Do not use ReturnDate here
            If boolGuWuAccess Then
                str1 = "SELECT * FROM TBLAUDITTRAIL WHERE DTSAVEDATE > #" & Format(dt1, "Short Date") & "# AND DTSAVEDATE <= #" & Format(dt2, "Short Date") & "#"
            ElseIf boolGuWuSQLServer Then
                str1 = "SELECT * FROM TBLAUDITTRAIL WHERE DTSAVEDATE > '" & Format(dt1, "Short Date") & "' AND DTSAVEDATE <= '" & Format(dt2, "Short Date") & "'"
            ElseIf boolGuWuOracle Then
                str1 = "SELECT * FROM TBLAUDITTRAIL WHERE DTSAVEDATE > TO_DATE('" & CStr(dt1) & "', 'YYYY/MM/DD') AND DTSAVEDATE <= TO_DATE('" & CStr(dt2) & "', 'YYYY/MM/DD')"
            End If

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
            Dim int1 As Int64
            int1 = rs1.RecordCount 'debug
            rs1.Close()
            rs1 = Nothing
            con.Close()
            con = Nothing
        Catch ex As Exception
            str1 = ex.Message
            str1 = str1
        End Try

        Cursor.Current = Cursors.Default

    End Sub

    Private Sub frmAuditTrail_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Call ControlDefaults(Me)

        'TODO: This line of code loads data into the 'GuWu_01DataSet.TBLAUDITTRAIL' table. You can move, or remove it, as needed.

        'position form

        'Dim ds As New DataSet
        Dim var1

        'Try
        '    ds.Tables.Add(tblAuditTrail)
        'Catch ex As Exception
        '    var1 = ex.Message
        '    var1 = var1
        'End Try

        'Me.TBLAUDITTRAILTableAdapter.ClearBeforeFill = True

        'default is last 3 months
        Dim dt1 As Date
        Dim dt2 As Date
        dt1 = Now
        dt2 = DateAdd(DateInterval.Day, -90, dt1)
        Me.DateTimePicker1.Value = dt2 ' CDate("1/1/2014")

        Call OpenConnection()

        Try
            Me.TBLAUDITTRAILBindingSource.DataSource = tblAuditTrail
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try
        'Me.TBLAUDITTRAILBindingSource.DataSource = ds

        Dim str1 As String
        str1 = "LABIntegrity StudyDoc" & ChrW(8482) & " Audit Trail"
        Me.Text = str1

        boolFormLoad = True

        Me.chkFilterDates.Checked = True

        Dim sw, sh
        sw = My.Computer.Screen.WorkingArea.Width
        sh = My.Computer.Screen.WorkingArea.Height
        Me.Left = sw * 0.05
        Me.Width = sw - (sw * 0.05 * 2)
        Me.Top = (sh - Me.Height) / 2

        Call FillSort()

        ' Me.TBLAUDITTRAILTableAdapter.ClearBeforeFill = True
        'Try
        '    Me.TBLAUDITTRAILTableAdapter.Fill(Me.GuWu_01DataSet.TBLAUDITTRAIL)
        'Catch ex As Exception
        '    var1 = ex.Message
        '    var1 = var1
        'End Try

        'var1 = Me.GuWu_01DataSet.TBLAUDITTRAIL.Rows.Count
        'var1 = var1

        'var1 = tblAuditTrail.Rows.Count
        'var1 = var1

        'Try
        '    Me.TBLAUDITTRAILTableAdapter.Fill(tblAuditTrail)
        'Catch ex As Exception
        '    var1 = ex.Message
        '    var1 = var1
        'End Try


        'Out
        'Me.dpAuditTrail.AllowUserToAddItems = False
        'Me.dpAuditTrail.AllowUserToDeleteItems = False

        'frmH.ta_tblReportStatements.Update(tblReportstatements)

        'Dim dv as system.data.dataview = New DataView(tblAuditTrail)

        Dim strF As String
        Dim strS As String

        'dv.AllowDelete = False
        'dv.AllowEdit = False
        'dv.AllowNew = False

        strF = "ID_TBLAUDITTRAIL < 0"
        strS = "ID_TBLAUDITTRAIL ASC"
        Me.TBLAUDITTRAILBindingNavigator.BindingSource.Filter = strF
        Me.TBLAUDITTRAILBindingNavigator.BindingSource.Sort = strS

        Me.DateTimePicker1.Format = DateTimePickerFormat.Short
        Me.DateTimePicker2.Format = DateTimePickerFormat.Short


        'Call FillRTF(strF, strS)

        'Me.lblPrint1.Visible = False

        'Out?
        'Me.rtfPrint.Left = Me.dpAuditTrail.Left
        'Me.rtfPrint.Top = Me.dpAuditTrail.Top
        'Me.rtfPrint.Width = Me.dpAuditTrail.Width
        'Me.rtfPrint.Height = Me.dpAuditTrail.Height

        'Me.dgvAT.Left = Me.dpAuditTrail.Left
        'Me.dgvAT.Top = Me.dpAuditTrail.Top
        'Me.dgvAT.Width = Me.dpAuditTrail.Width
        'Me.dgvAT.Height = Me.dpAuditTrail.Height

        boolFormLoad = False

        Call FillDP()

        If boolFromRW Then
            Call DoRW()
        End If

    End Sub

    Sub DoRW()

        Me.cbxFilterSource.SelectedIndex = 0

        'find study in cbxFilter1
        Dim str1 As String
        Dim str2 As String
        Dim Count1 As Int64

        str1 = frmH.cbxStudy.Text

        Dim int1 As Int64

        int1 = Me.cbxFilter1.FindStringExact(str1)

        If int1 = -1 Then
        Else
            Me.cbxFilter1.SelectedIndex = int1
        End If


    End Sub

    Sub FillDP()

        If boolFormLoad Then
            Exit Sub
        End If

        If boolHold Then
            Exit Sub
        End If

        Dim strF As String
        Dim strF1 As String
        Dim strFC As String
        Dim strS As String
        Dim dt1 As Date
        Dim dt2 As Date
        Dim strASC As String
        Dim var1
        Dim ID As Int64

        Cursor.Current = Cursors.WaitCursor

        Try

            strFC = Me.cbxFilter1.Text
            Select Case Me.cbxFilterSource.Text
                Case "Study Name"
                    If Len(strFC) = 0 Then
                        strF = "ID_TBLAUDITTRAIL < 0"
                    Else
                        strF = "CHARWATSONSTUDYNAME = '" & Me.cbxFilter1.Text & "'"
                        strF = "ID_TBLSTUDIES = '" & Me.cbxFilter1.SelectedValue & "'"
                    End If
                Case "User Name"
                    If Len(strFC) = 0 Then
                        strF = "ID_TBLAUDITTRAIL < 0"
                    Else
                        'find
                        strF = "CHARUSERNAME = '" & strFC & "'"
                    End If
                Case "StudyDoc Administration"
                    strF = "CHARAUDITTYPE = 'StudyDoc Administration'"
                Case "Report Writer Administration"
                    strF = "CHARAUDITTYPE = 'Report Writer Administration'"

                Case Else
                    If InStr(1, Me.cbxFilterSource.Text, "Word Template", CompareMethod.Text) > 0 Then
                        'need to find ID_TBLWORDSTATEMENTS in TBLWORDSTATEMENTS
                        'strF = "(CHARNEWVALUE = '" & strFC & "' AND CHARTABLE = 'TBLWORDSTATEMENTS') OR (CHARTABLE = 'TBLWORDDOCS' AND )"
                        strF = "CHARTITLE = '" & Me.cbxFilter1.Text & "'"
                        Dim rows() As DataRow
                        rows = tblWordStatements.Select(strF)
                        If rows.Length = 0 Then
                            ID = 0
                        Else
                            ID = rows(0).Item("ID_TBLWORDSTATEMENTS")
                        End If

                        'strF = "(CHARTABLE LIKE 'TBLWORDSTATEMENTS*') AND (CHARLINK2VALUE = '" & ID.ToString & "')"
                        strF = "(CHARTABLE LIKE 'TBLWORDSTATEMENTS*') AND (ID_SOURCETABLE = '" & ID.ToString & "')"
                    Else
                        strF = "ID_TBLAUDITTRAIL < 0"
                    End If

            End Select

            'now check for dates
            If Me.chkFilterDates.Checked Then

                dt1 = CDate(Format(Me.DateTimePicker1.Value, "yyyy/MM/dd"))
                dt2 = CDate(Format(Me.DateTimePicker2.Value, "yyyy/MM/dd"))

                If dt2 < dt1 Then 'ignore
                Else
                    strF1 = "DTSAVEDATE >= " & ReturnDate(dt1) & " AND DTSAVEDATE <= " & ReturnDate(DateAdd(DateInterval.Day, 1, dt2))

                    'If boolGuWuAccess Then
                    '    strF1 = "DTSAVEDATE > #" & dt1 & "# AND DTSAVEDATE <= #" & DateAdd(DateInterval.Day, 1, dt2) & "#"
                    'ElseIf boolGuWuSQLServer Then
                    '    strF1 = "DTSAVEDATE > '" & dt1 & "' AND DTSAVEDATE <= '" & DateAdd(DateInterval.Day, 1, dt2) & "'"
                    'ElseIf boolGuWuOracle Then
                    '    strF1 = "DTSAVEDATE > TO_DATE('" & CStr(dt1) & "', 'YYYY/MM/DD') AND DTSAVEDATE <= TO_DATE('" & CStr(DateAdd(DateInterval.Day, 1, dt2)) & "', 'YYYY/MM/DD')"
                    'End If

                    strF = strF & " AND " & strF1
                End If

            End If

            'now do sort
            Dim S1 As String
            Dim S2 As String
            Dim S3 As String
            Dim S4 As String
            Dim strCol As String

            S1 = Me.cbxSort1.Text
            S2 = Me.cbxSort2.Text
            S3 = Me.cbxSort3.Text
            S4 = Me.cbxSort4.Text

            strS = ""
            If Len(S1) = 0 Then
            Else
                strASC = "ASC"
                If Me.rbSort1Asc.Checked Then
                    strASC = "ASC"
                Else
                    strASC = "DESC"
                End If
                Select Case S1
                    Case "Study Name"
                        strCol = "CHARWATSONSTUDYNAME"
                    Case "User Name"
                        strCol = "CHARUSERNAME"
                    Case "Changed Item"
                        strCol = "CHARACTUALITEM"
                    Case "Event Date"
                        strCol = "DTSAVEDATE"
                End Select
                If Len(strS) = 0 Then
                    strS = strCol & " " & strASC
                Else
                    strS = strS & ", " & strCol & " " & strASC
                End If
            End If

            If Len(S2) = 0 Then
            Else
                strASC = "ASC"
                If Me.rbSort2Asc.Checked Then
                    strASC = "ASC"
                Else
                    strASC = "DESC"
                End If
                Select Case S2
                    Case "Study Name"
                        strCol = "CHARWATSONSTUDYNAME"
                    Case "User Name"
                        strCol = "CHARUSERNAME"
                    Case "Changed Item"
                        strCol = "CHARACTUALITEM"
                    Case "Event Date"
                        strCol = "DTSAVEDATE"
                End Select
                If Len(strS) = 0 Then
                    strS = strCol & " " & strASC
                Else
                    strS = strS & ", " & strCol & " " & strASC
                End If
            End If

            If Len(S3) = 0 Then
            Else
                strASC = "ASC"
                If Me.rbSort3Asc.Checked Then
                    strASC = "ASC"
                Else
                    strASC = "DESC"
                End If
                Select Case S3
                    Case "Study Name"
                        strCol = "CHARWATSONSTUDYNAME"
                    Case "User Name"
                        strCol = "CHARUSERNAME"
                    Case "Changed Item"
                        strCol = "CHARACTUALITEM"
                    Case "Event Date"
                        strCol = "DTSAVEDATE"
                End Select
                If Len(strS) = 0 Then
                    strS = strCol & " " & strASC
                Else
                    strS = strS & ", " & strCol & " " & strASC
                End If
            End If

            If Len(S4) = 0 Then
            Else
                strASC = "ASC"
                If Me.rbSort4Asc.Checked Then
                    strASC = "ASC"
                Else
                    strASC = "DESC"
                End If
                Select Case S4
                    Case "Study Name"
                        strCol = "CHARWATSONSTUDYNAME"
                    Case "User Name"
                        strCol = "CHARUSERNAME"
                    Case "Changed Item"
                        strCol = "CHARACTUALITEM"
                    Case "Event Date"
                        strCol = "DTSAVEDATE"
                End Select
                If Len(strS) = 0 Then
                    strS = strCol & " " & strASC
                Else
                    strS = strS & ", " & strCol & " " & strASC
                End If
            End If


            Me.TBLAUDITTRAILBindingNavigator.BindingSource.Filter = strF
            Me.TBLAUDITTRAILBindingNavigator.BindingSource.Sort = strS

            Me.TBLAUDITTRAILBindingNavigator.Refresh()

            Me.Refresh()

            Dim dgv As DataGridView
            Dim dv As system.data.dataview = New DataView(tblAuditTrail)
            Dim Count1 As Short
            Dim str2 As String

            var1 = dv.Count 'debug

            dv.RowFilter = strF
            dv.Sort = strS
            dv.AllowDelete = False
            dv.AllowEdit = False
            dv.AllowNew = False

            dgv = Me.dgvAT
            dgv.ReadOnly = True
            dgv.DataSource = dv

            var1 = dv.Count

            For count1 = 0 To dgv.Columns.Count - 1
                dgv.Columns(Count1).Visible = False
            Next

            dgv.Columns("CHARUSERNAME").HeaderText = "User Name"
            dgv.Columns("CHARUSERNAME").Visible = True
            dgv.Columns("CHARUSERNAME").DisplayIndex = 7

            dgv.Columns("DTSAVEDATE").HeaderText = "Event Date/Time"
            dgv.Columns("DTSAVEDATE").Visible = True
            dgv.Columns("DTSAVEDATE").DisplayIndex = 6
            str2 = "MMM dd, yyyy  HH:mm:ss tt"
            'str2 = "MMM dd, yyyy HH:mm:ss"
            dgv.Columns("DTSAVEDATE").DefaultCellStyle.Format = str2

            dgv.Columns("CHARNEWVALUE").HeaderText = "New Value"
            dgv.Columns("CHARNEWVALUE").Visible = True
            dgv.Columns("CHARNEWVALUE").DisplayIndex = 5
            dgv.Columns("CHARNEWVALUE").DefaultCellStyle.WrapMode = DataGridViewTriState.True

            dgv.Columns("CHAROLDVALUE").HeaderText = "Old Value"
            dgv.Columns("CHAROLDVALUE").Visible = True
            dgv.Columns("CHAROLDVALUE").DisplayIndex = 4
            dgv.Columns("CHAROLDVALUE").DefaultCellStyle.WrapMode = DataGridViewTriState.True

            dgv.Columns("CHARACTUALITEM").HeaderText = "Item Changed"
            dgv.Columns("CHARACTUALITEM").Visible = True
            dgv.Columns("CHARACTUALITEM").DisplayIndex = 3

            dgv.Columns("CHARACTION").HeaderText = "Action"
            dgv.Columns("CHARACTION").Visible = True
            dgv.Columns("CHARACTION").DisplayIndex = 2
            dgv.Columns("CHARACTION").DefaultCellStyle.WrapMode = DataGridViewTriState.True

            dgv.Columns("CHARTABLEDESCRIPTION").HeaderText = "StudyDoc Source"
            dgv.Columns("CHARTABLEDESCRIPTION").Visible = True
            dgv.Columns("CHARTABLEDESCRIPTION").DisplayIndex = 1

            dgv.Columns("CHARWATSONSTUDYNAME").HeaderText = "Study Name"
            dgv.Columns("CHARWATSONSTUDYNAME").Visible = True
            dgv.Columns("CHARWATSONSTUDYNAME").DisplayIndex = 0

            'do some again

            dgv.Columns("CHARNEWVALUE").DisplayIndex = 5
            dgv.Columns("CHAROLDVALUE").DisplayIndex = 4
            dgv.Columns("CHARACTUALITEM").DisplayIndex = 3
            dgv.Columns("CHARACTION").DisplayIndex = 2
            dgv.Columns("CHARWATSONSTUDYNAME").DisplayIndex = 0

            dgv.Columns("CHARWATSONSTUDYNAME").DisplayIndex = 0
            dgv.Columns("CHARACTION").DisplayIndex = 2
            dgv.Columns("CHARACTUALITEM").DisplayIndex = 3
            dgv.Columns("CHAROLDVALUE").DisplayIndex = 4
            dgv.Columns("CHARNEWVALUE").DisplayIndex = 5

            'dgv.Columns("CHARNEWVALUE").DisplayIndex = 5
            'dgv.Columns("CHAROLDVALUE").DisplayIndex = 4
            'dgv.Columns("CHARACTUALITEM").DisplayIndex = 3
            'dgv.Columns("CHARACTION").DisplayIndex = 2
            'dgv.Columns("CHARWATSONSTUDYNAME").DisplayIndex = 0

            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders
            dgv.AutoResizeColumns()
            dgv.AutoResizeRows()

            If Me.rbPrint.Checked Then
                Call FillRTF(strF, strS)
            End If

        Catch ex As Exception

            var1 = ex.Message
            var1 = var1

        End Try

        Cursor.Current = Cursors.Default

        boolNeedUpdate = True

    End Sub


    Private Sub PrintDocument1_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage

        '' Print the content of the RichTextBox. Store the last character printed.
        'checkPrint = Me.rtfPrint.Print(checkPrint, Me.rtfPrint.TextLength, e)

        '' Look for more pages
        'If checkPrint < Me.rtfPrint.TextLength Then
        '    e.HasMorePages = True
        'Else
        '    e.HasMorePages = False
        'End If

    End Sub


    Sub FillFilter1()


        Dim boolStudy As Boolean = False
        Dim boolUserName As Boolean = False
        Dim dtbl As System.Data.DataTable
        Dim dv As system.data.dataview
        Dim strMember As String
        Dim strF As String
        Dim strS As String
        Dim Count1 As Int64
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim int1 As Int32
        Dim id1 As Int64
        Dim strFF As String

        Select Case Me.cbxFilterSource.Text
            Case "Study Name"
                boolStudy = True
            Case "User Name"
                boolUserName = True
            Case Else

        End Select

        strFF = Me.cbxFilterSource.Text

        Select Case strFF
            Case "Study Name"

                '20171113 LEE:
                'User can now delete study
                'Must get fill cbxFilter1 with unique query from tblAuditTrail
                strS = "CHARWATSONSTUDYNAME ASC, ID_TBLSTUDIES ASC, ID_TBLAUDITTRAIL ASC"
                Dim dvAT As DataView = New DataView(tblAuditTrail, "ID_TBLSTUDIES <> 0", strS, DataViewRowState.CurrentRows)
                Dim tblAT As DataTable = dvAT.ToTable("a", True, "ID_TBLSTUDIES", "CHARWATSONSTUDYNAME")

                dtbl = tblStudies
                strF = "ID_TBLSTUDIES <> 0"
                strS = "" '"CHARWATSONSTUDYNAME ASC, ID_TBLSTUDIES ASC"
                'sorted earlier
                dv = New DataView(tblAT, strF, strS, DataViewRowState.CurrentRows)
                strMember = "ID_TBLSTUDIES" ' "CHARWATSONSTUDYNAME"
                Me.cbxFilter1.DataSource = dv
                Me.cbxFilter1.DisplayMember = "CHARWATSONSTUDYNAME"
                Me.cbxFilter1.ValueMember = "ID_TBLSTUDIES"

            Case "User Name"

                Me.cbxFilter1.DataSource = Nothing
                Me.cbxFilter1.Items.Clear()

                dtbl = tblPersonnel
                strF = "ID_TBLSTUDIES > 0"
                strS = "CHARWATSONSTUDYNAME ASC"

                Dim tbl2 As System.Data.DataTable
                Dim rows2() As DataRow
                tbl2 = tblUserAccounts

                For Count1 = 0 To dtbl.Rows.Count - 1

                    id1 = dtbl.Rows(Count1).Item("ID_TBLPERSONNEL")
                    strF = "ID_TBLPERSONNEL = " & id1
                    Erase rows2
                    rows2 = tbl2.Select(strF)
                    If rows2.Length > 0 Then
                        str1 = NZ(dtbl.Rows(Count1).Item("charFIRSTNAME"), "NA")
                        str2 = NZ(dtbl.Rows(Count1).Item("charMIDDLEname"), "")
                        str3 = NZ(dtbl.Rows(Count1).Item("charLASTNAME"), "NA")
                        'If StrComp(str3, "aaAdmin", CompareMethod.Text) = 0 Or StrComp(str3, "Elvebak", CompareMethod.Text) = 0 Then
                        If StrComp(str3, "Elvebak", CompareMethod.Text) = 0 Then
                        Else
                            If Len(str2) = 0 Then 'no middle initial provided
                                str4 = str1 & " " & str3
                            Else
                                'If Len(str2) = 1 Then 'needs a period
                                '    str2 = str2 & "."
                                'Else
                                'End If
                                str4 = str1 & " " & str2 & " " & str3
                                'str4 = str3 & ", " & str1 & " " & str2
                            End If
                            Me.cbxFilter1.Items.Add(str4)
                        End If
                    End If

                Next

            Case "StudyDoc Administration"
                Me.cbxFilter1.DataSource = Nothing
                Me.cbxFilter1.Items.Clear()


            Case "Report Writer Administration"
                Me.cbxFilter1.DataSource = Nothing
                Me.cbxFilter1.Items.Clear()

            Case Else
                If InStr(1, strFF, "Word Template", CompareMethod.Text) > 0 Then
                    'fill cbxFilter1 with Worddoc names
                    dtbl = tblWordStatements
                    strF = "ID_TBLWORDSTATEMENTS > 0"
                    strS = "CHARTITLE ASC"
                    dv = New DataView(dtbl, strF, strS, DataViewRowState.CurrentRows)
                    strMember = "CHARTITLE"
                    Me.cbxFilter1.DataSource = dv
                    Me.cbxFilter1.DisplayMember = strMember

                End If


        End Select

        Me.cbxFilter1.SelectedIndex = -1


    End Sub

    Sub FillRTF(ByVal strF As String, ByVal strS As String)

        If boolNeedUpdate Then

            Me.rtfPrint.Left = Me.pan1.Left
            Me.rtfPrint.Top = Me.pan1.Top
            Me.rtfPrint.Width = (Me.pan2.Left + Me.pan2.Width) - Me.pan1.Left
            Me.rtfPrint.Height = Me.pan1.Height

            Me.dgvAT.Left = Me.rtfPrint.Left
            Me.dgvAT.Top = Me.rtfPrint.Top
            Me.dgvAT.Width = Me.rtfPrint.Width
            Me.dgvAT.Height = Me.rtfPrint.Height

            Me.rtfPrint.AcceptsTab = True
            Me.rtfPrint.SelectionIndent = 10 ' 100
            Me.rtfPrint.SelectionRightIndent = 25
            Me.rtfPrint.SelectionHangingIndent = 300

            'Me.rtfPrint.SelectionTabs = New Integer() {100, 80, 120, 160}
            Me.rtfPrint.SelectionTabs = New Integer() {300}

            'clear rtf
            Me.rtfPrint.Text = ""
            'now load rtf
            Me.rtfPrint.WordWrap = True

            Dim Count1 As Int64
            Dim dv As System.Data.DataView = New DataView(tblAuditTrail)

            dv.RowFilter = strF
            dv.Sort = strS

            Dim strAT As String
            Dim strHeader As String
            'now do sort
            Dim S1 As String
            Dim S2 As String
            Dim S3 As String
            Dim S4 As String
            Dim strSort As String
            Dim strDt1 As String
            Dim strDt2 As String

            Dim strDF As String = "MMM dd, yyyy  HH:mm:ss tt"

            strDt1 = "Not Applicable"
            strDt2 = "Not Applicable"

            Try
                strDt1 = Format(Me.DateTimePicker1.Value, "MMM dd, yyyy")
            Catch ex As Exception
                strDt1 = "Not Applicable"
            End Try

            Try
                strDt2 = Format(Me.DateTimePicker2.Value, "MMM dd, yyyy")
            Catch ex As Exception
                strDt2 = "Not Applicable"
            End Try

            S1 = Me.cbxSort1.Text
            S2 = Me.cbxSort2.Text
            S3 = Me.cbxSort3.Text
            S4 = Me.cbxSort4.Text

            strSort = ""
            Dim strASC As String
            Dim strCol As String

            'set text
            If Len(S1) = 0 Then
            Else
                strASC = "ASC"
                If Me.rbSort1Asc.Checked Then
                    strASC = "ASC"
                Else
                    strASC = "DESC"
                End If
                If Len(strSort) = 0 Then
                    strSort = Me.cbxSort1.Text & " " & strASC
                Else
                    strSort = strSort & ", " & Me.cbxSort1.Text & " " & strASC
                End If
            End If

            If Len(S2) = 0 Then
            Else
                strASC = "ASC"
                If Me.rbSort2Asc.Checked Then
                    strASC = "ASC"
                Else
                    strASC = "DESC"
                End If
                If Len(strSort) = 0 Then
                    strSort = Me.cbxSort2.Text & " " & strASC
                Else
                    strSort = strSort & ", " & Me.cbxSort2.Text & " " & strASC
                End If
            End If

            If Len(S3) = 0 Then
            Else
                strASC = "ASC"
                If Me.rbSort3Asc.Checked Then
                    strASC = "ASC"
                Else
                    strASC = "DESC"
                End If
                If Len(strSort) = 0 Then
                    strSort = Me.cbxSort3.Text & " " & strASC
                Else
                    strSort = strSort & ", " & Me.cbxSort3.Text & " " & strASC
                End If
            End If

            If Len(S4) = 0 Then
            Else
                strASC = "ASC"
                If Me.rbSort4Asc.Checked Then
                    strASC = "ASC"
                Else
                    strASC = "DESC"
                End If
                If Len(strSort) = 0 Then
                    strSort = Me.cbxSort4.Text & " " & strASC
                Else
                    strSort = strSort & ", " & Me.cbxSort4.Text & " " & strASC
                End If
            End If

            Dim strAT1 As String

            strAT = ""
            strAT1 = ""

            Me.lblStatus.Text = "Filling Print View: 0 of " & dv.Count
            If dv.Count > 100 Then
                Me.lblStatus.Visible = True
                Me.lblStatus.Refresh()
            End If

            For Count1 = 0 To dv.Count - 1

                Me.lblStatus.Text = "Filling Print View: " & Count1 + 1 & " of  " & dv.Count ' & " (" & Len(strAT) & ")"
                Me.lblStatus.Refresh()

                strHeader = "LABIntegrity StudyDoc" & ChrW(8482) & " Audit Trail Report" & ChrW(10)
                'strHeader = strHeader & "Record:" & ChrW(9) & Count1 + 1 & " of  " & dv.Count & ChrW(10)
                strHeader = strHeader & "Report Printed Date:" & ChrW(9) & "Print Not Applicable" & ChrW(10)
                strHeader = strHeader & "Report Printed By:" & ChrW(9) & gUserName & " (User ID:  " & gUserID & ")" & ChrW(10)
                strHeader = strHeader & ChrW(10)

                Select Case Me.cbxFilterSource.Text
                    Case "Study Name"
                        strHeader = strHeader & "Audit Trail Filtered by:" & ChrW(9) & "Study Name:" & ChrW(9) & Me.cbxFilter1.Text & ChrW(10)
                    Case "User Name"
                        strHeader = strHeader & "Audit Trail Filtered by:" & ChrW(9) & "User Name:" & ChrW(9) & Me.cbxFilter1.Text & ChrW(10)
                    Case "StudyDoc Administration"
                        strHeader = strHeader & "Audit Trail Filtered by:" & ChrW(9) & "StudyDoc Administration items" & ChrW(10)
                    Case "Report Writer Administration"
                        strHeader = strHeader & "Audit Trail Filtered by:" & ChrW(9) & "Report Writer Administration items" & ChrW(10)
                    Case Else
                        strHeader = strHeader & "Audit Trail Filtered by:" & ChrW(9) & "Not Applicable" & ChrW(10)
                End Select

                If Me.chkFilterDates.Checked Then
                    strHeader = strHeader & "Audit Trail Dates Filtered between:" & ChrW(9) & strDt1 & " and " & strDt2 & ChrW(10)
                Else
                    strHeader = strHeader & "Audit Trail Dates Filtered between:" & ChrW(9) & "Not Applicable" & ChrW(10)
                End If
                If Len(strSort) = 0 Then
                    strHeader = strHeader & "Audit Trail Sorted by:" & ChrW(9) & "Not Applicable" & ChrW(10)
                Else
                    strHeader = strHeader & "Audit Trail Sorted by:" & ChrW(9) & strSort & ChrW(10)
                End If
                strHeader = strHeader & ChrW(10)

                'If Count1 = 0 Then
                '    strAT = strHeader & "StudyDoc Audit Trail Report Record:" & ChrW(9) & Count1 + 1 & " of  " & dv.Count & ChrW(10)
                'Else
                '    strAT = strAT & strHeader & "StudyDoc Audit Trail Record:" & ChrW(9) & Count1 + 1 & " of  " & dv.Count & ChrW(10)
                'End If

                strAT = strAT & strHeader & "StudyDoc Audit Trail Record:" & ChrW(9) & Count1 + 1 & " of  " & dv.Count & ChrW(10)
                strAT = strAT & ChrW(10)
                strAT = strAT & "StudyDoc Source:" & ChrW(9) & dv(Count1).Item("CHARTABLEDESCRIPTION") & ChrW(10)
                strAT = strAT & "Where:" & ChrW(10)
                strAT = strAT & "       " & dv(Count1).Item("CHARLINK1") & " =" & ChrW(9) & dv(Count1).Item("CHARLINK1VALUE") & ChrW(10)
                strAT = strAT & "and" & ChrW(10)
                strAT = strAT & "       " & dv(Count1).Item("CHARLINK2") & " =" & ChrW(9) & dv(Count1).Item("CHARLINK2VALUE") & ChrW(10)
                strAT = strAT & "Changed Item:" & ChrW(9) & dv(Count1).Item("CHARACTUALITEM") & ChrW(10)
                strAT = strAT & "Action:" & ChrW(9) & dv(Count1).Item("CHARACTION") & ChrW(10)
                strAT = strAT & "Old Value:" & ChrW(9) & dv(Count1).Item("CHAROLDVALUE") & ChrW(10)
                strAT = strAT & "New Value:" & ChrW(9) & dv(Count1).Item("CHARNEWVALUE") & ChrW(10)
                'strAT = strAT & ChrW(10)
                strAT = strAT & "StudyDoc Study ID:" & ChrW(9) & dv(Count1).Item("ID_TBLSTUDIES") & ChrW(10)
                strAT = strAT & "Study Name:" & ChrW(9) & dv(Count1).Item("CHARWATSONSTUDYNAME") & ChrW(10)
                strAT = strAT & "Event Date/Time:" & ChrW(9) & Format(CDate(dv(Count1).Item("DTSAVEDATE")), strDF) & ChrW(10)
                strAT = strAT & "Standard Time Zone Name:" & ChrW(9) & dv(Count1).Item("CHARSTANDARDTIMEZONE") & ChrW(10)
                strAT = strAT & "Daylight Saving Time Name:" & ChrW(9) & dv(Count1).Item("CHARDAYLIGHTSAVINGZONE") & ChrW(10)
                strAT = strAT & "Daylight Saving Time?:" & ChrW(9) & dv(Count1).Item("CHARDAYLIGHTSAVINGTIME") & ChrW(10)
                strAT = strAT & "Coordinated Universal Time:" & ChrW(9) & dv(Count1).Item("CHARCOORUNIVTIME") & ChrW(10)
                strAT = strAT & "UTC Offset:" & ChrW(9) & dv(Count1).Item("CHARUTCOFFSET") & ChrW(10)
                strAT = strAT & "Workstation Name:" & ChrW(9) & dv(Count1).Item("CHARWORKSTATION") & ChrW(10)
                strAT = strAT & "User Name:" & ChrW(9) & dv(Count1).Item("CHARUSERNAME") & ChrW(10)
                strAT = strAT & "User ID:" & ChrW(9) & dv(Count1).Item("CHARUSERID") & ChrW(10)
                strAT = strAT & "Reason for Change:" & ChrW(9) & dv(Count1).Item("CHARTBLREASONFORCHANGE") & ChrW(10)
                strAT = strAT & "Meaning of Signature:" & ChrW(9) & dv(Count1).Item("CHARTBLCHARMEANINGOFSIG") & ChrW(10)
                'strAT = strAT & ChrW(10)
                strAT = strAT & "Database Table:" & ChrW(9) & dv(Count1).Item("CHARTABLE") & ChrW(10)
                strAT = strAT & "Database Table Column:" & ChrW(9) & dv(Count1).Item("CHARCOLUMN") & ChrW(10)
                strAT = strAT & "Database Table Record ID:" & ChrW(9) & dv(Count1).Item("ID_SOURCETABLE") & ChrW(10)
                strAT = strAT & "End of Record"
                If Count1 = dv.Count - 1 Then
                    strAT = strAT & ChrW(10)
                Else
                    strAT = strAT & ChrW(10) & ChrW(12) & ChrW(10)
                End If

                If Len(strAT) > 4000 Then
                    strAT1 = strAT1 & strAT
                    strAT = ""
                End If

            Next

            strAT1 = strAT1 & strAT

            Me.rtfPrint.Text = strAT1

            boolNeedUpdate = False

        End If

        Me.lblStatus.Visible = False

    End Sub

    Sub DoView()

        Call frmResize()

        Cursor.Current = Cursors.WaitCursor

        If Me.rbLong.Checked Then

            Me.pan1.Visible = True
            Me.pan2.Visible = True
            Me.rtfPrint.Visible = False
            Me.dgvAT.Visible = False
            Me.cmdClipboard.Visible = False

        ElseIf Me.rbShort.Checked Then

            Me.dgvAT.Visible = True
            Me.rtfPrint.Visible = False
            Me.pan1.Visible = False
            Me.pan2.Visible = False
            Me.cmdClipboard.Visible = True

        ElseIf Me.rbPrint.Checked Then

            Me.rtfPrint.Visible = True
            Me.pan1.Visible = False
            Me.pan2.Visible = False
            Me.dgvAT.Visible = False
            Me.cmdClipboard.Visible = False

            Dim strF As String
            Dim strS As String

            strF = Me.TBLAUDITTRAILBindingNavigator.BindingSource.Filter ' = strF
            strS = Me.TBLAUDITTRAILBindingNavigator.BindingSource.Sort ' = strS

            Call FillRTF(strF, strS)

        End If

        Cursor.Current = Cursors.Default

    End Sub


    Private Sub cmdRTF_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRTF.Click

        Call frmResize()

        Cursor.Current = Cursors.WaitCursor

        If InStr(1, Me.cmdRTF.Text, "Show", CompareMethod.Text) > 0 Then
            Me.cmdRTF.Text = "Hide Print Version"
            Me.cmdShort.Text = "Show Short Version"
            Me.rtfPrint.Visible = True
            Me.pan1.Visible = False
            Me.pan2.Visible = False
            Me.dgvAT.VirtualMode = False

            Dim strF As String
            Dim strS As String

            strF = Me.TBLAUDITTRAILBindingNavigator.BindingSource.Filter ' = strF
            strS = Me.TBLAUDITTRAILBindingNavigator.BindingSource.Sort ' = strS

            Call FillRTF(strF, strS)

            Me.cmdPrint.Text = "Show Short Version"

        Else
            Me.cmdRTF.Text = "Show Print Version"
            Me.cmdShort.Text = "Show Short Version"
            Me.rtfPrint.Visible = False
            Me.pan1.Visible = True
            Me.pan2.Visible = True
            Me.dgvAT.Visible = False

        End If

        Cursor.Current = Cursors.Default

    End Sub

    Private Sub chkFilterDates_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFilterDates.CheckedChanged

        'Call DoDateFilters()

        'If boolFormLoad Then
        '    Exit Sub
        'End If

        'Call FillDP()


    End Sub

    Private Sub cbxFilter1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxFilter1.SelectedIndexChanged

        Call FillDP()

    End Sub

    Private Sub DateTimePicker1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles DateTimePicker1.Validating

        ''must be <= to dtpicker2
        'Dim dt1 As Date
        'Dim dt2 As Date
        'Dim strM As String

        'dt1 = CDate(Format(Me.DateTimePicker1.Value, "MMM dd, yyyy"))
        'dt2 = CDate(Format(Me.DateTimePicker2.Value, "MMM dd, yyyy"))

        'strM = "First date must <= Last date"
        'If dt1 > dt2 Then
        '    MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
        '    e.Cancel = True
        'Else
        '    Call FillDP()
        'End If


    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged

        ''must be <= to dtpicker2
        'Dim dt1 As Date
        'Dim dt2 As Date
        'Dim strM As String

        'dt1 = CDate(Format(Me.DateTimePicker1.Value, "MMM dd, yyyy"))
        'dt2 = CDate(Format(Me.DateTimePicker2.Value, "MMM dd, yyyy"))

        'strM = "First date must <= Last date"
        'If dt1 > dt2 Then
        '    MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
        'Else
        '    Call FillDP()
        'End If


    End Sub


    Private Sub DateTimePicker2_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles DateTimePicker2.Validating

        ''must be >= to dtpicker2
        'Dim dt1 As Date
        'Dim dt2 As Date
        'Dim strM As String

        'dt1 = CDate(Format(Me.DateTimePicker1.Value, "MMM dd, yyyy"))
        'dt2 = CDate(Format(Me.DateTimePicker2.Value, "MMM dd, yyyy"))

        'strM = "Last date must >= First date"
        'If dt2 < dt1 Then
        '    MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
        '    e.Cancel = True
        'End If

    End Sub

    Private Sub DateTimePicker2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker2.ValueChanged




    End Sub

    Private Sub cmdDateUpdate_Click(sender As System.Object, e As System.EventArgs) Handles cmdDateUpdate.Click

        'must be >= to dtpicker2
        Dim dt1 As Date
        Dim dt2 As Date
        Dim strM As String

        Call OpenConnection()

        dt1 = CDate(Format(Me.DateTimePicker1.Value, "MMM dd, yyyy"))
        dt2 = CDate(Format(Me.DateTimePicker2.Value, "MMM dd, yyyy"))

        strM = "Last date must >= First date"
        If dt2 < dt1 Then
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
        Else
            Call FillDP()
        End If

    End Sub

    Private Sub cbxSort1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSort1.SelectedIndexChanged

        Call FillDP()

    End Sub

    Private Sub cbxSort2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSort2.SelectedIndexChanged

        Call FillDP()

    End Sub

    Private Sub cbxSort3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSort3.SelectedIndexChanged

        Call FillDP()

    End Sub

    Private Sub cbxSort4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSort4.SelectedIndexChanged

        Call FillDP()

    End Sub

    Private Sub cmdPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdPrint.Click

        Call frmResize()

        Dim strF As String
        Dim strS As String
        Dim int1 As Int32

        Dim localZone As TimeZone = TimeZone.CurrentTimeZone
        Dim strTimeZoneName As String
        strTimeZoneName = NZ(localZone.StandardName, "NA")

        If Me.dgvAT.Rows.Count = 0 Then
            strS = "Nothing to print"
            MsgBox(strS, MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If

        If boolNeedUpdate Then

            strF = Me.TBLAUDITTRAILBindingNavigator.BindingSource.Filter ' = strF
            strS = Me.TBLAUDITTRAILBindingNavigator.BindingSource.Sort ' = strS

            Call FillRTF(strF, strS)

        End If

        'If Me.PrintDialog1.ShowDialog() = DialogResult.OK Then
        '    PrintDocument1.Print()
        'End If

        Dim frm As New frmAuditTrailPrint

        Call frm.FormatThis()

        Dim strPrint As String
        Dim strDt As String

        strDt = Format(Now, "MMM dd, yyyy HH:mm:ss tt") & " " & strTimeZoneName

        strPrint = Replace(Me.rtfPrint.Text, "Print Not Applicable", strDt, 1, -1, CompareMethod.Text)

        frm.rtfPrint.Text = strPrint

        int1 = CLng(Me.BindingNavigatorPositionItem.Text)

        frm.intSelPage = int1

        'frm.ShowDialog() 'debug

        Call frm.PrintThis()

        frm.Dispose()


    End Sub

    Private Sub rbSort1Asc_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSort1Asc.CheckedChanged

        Call FillDP()

    End Sub


    Private Sub rbSort2Asc_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSort2Asc.CheckedChanged

        Call FillDP()

    End Sub


    Private Sub rbSort3Asc_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSort3Asc.CheckedChanged

        Call FillDP()

    End Sub


    Private Sub rbSort4Asc_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSort4Asc.CheckedChanged

        Call FillDP()

    End Sub


    Private Sub cbxFilterSource_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxFilterSource.SelectedIndexChanged

        boolHold = True
        Call FillFilter1()
        boolHold = False

        Call FillDP()

    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click

        Me.Dispose()

    End Sub

    Private Sub cmdShort_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdShort.Click

        Call frmResize()

        Dim str1 As String

        str1 = Me.cmdShort.Text

        Cursor.Current = Cursors.WaitCursor

        Me.rtfPrint.Visible = False
        Me.cmdRTF.Text = "Show Print Version"

        If InStr(1, str1, "Long", CompareMethod.Text) > 0 Then

            Me.pan1.Visible = True
            Me.pan2.Visible = True
            Me.dgvAT.Visible = False
            Me.cmdShort.Text = "Show Short Version"
        Else

            Me.pan1.Visible = False
            Me.pan2.Visible = False
            Me.dgvAT.Visible = True
            Me.dgvAT.AutoResizeColumns()
            Me.cmdShort.Text = "Show Long Version"
        End If

        Cursor.Current = Cursors.Default

    End Sub

    Sub frmResize()

        Dim fw, p1w, p2w
        Dim a, b, c, d

        fw = Me.Width

        Me.pan2.Width = fw / 2 * 0.97
        Me.pan1.Width = Me.pan2.Width
        Me.pan2.Left = Me.pan1.Width + Me.pan1.Left + 5

        Me.pan2a.Top = 3
        Me.pan2a.Left = 3
        'Me.pan2a.Height = Me.pan2.Height - 6
        Me.pan2a.Width = Me.pan2.Width - 6

        Me.ID_TBLSTUDIESTextBox.Anchor = AnchorStyles.Left + AnchorStyles.Right + AnchorStyles.Top

        Me.pan1a.Top = 3
        Me.pan1a.Left = 3
        'Me.pan1a.Height = Me.pan1.Height - 6
        Me.pan1a.Width = Me.pan1.Width - 6

        Me.gbxlabelAuditTrail1.Top = Me.pan2.Top - Me.gbxlabelAuditTrail1.Height
        Me.gbxlabelAuditTrail1.Left = Me.pan2.Left

        Me.rtfPrint.Left = Me.pan1.Left + Me.pan1a.Left
        Me.rtfPrint.Top = Me.pan1.Top + Me.pan1a.Top
        Me.rtfPrint.Width = (Me.pan2.Left + Me.pan2.Width) - Me.pan1a.Left
        Me.rtfPrint.Height = Me.pan1.Height
        Me.rtfPrint.Refresh()

        Me.dgvAT.Left = Me.rtfPrint.Left
        Me.dgvAT.Top = Me.rtfPrint.Top
        Me.dgvAT.Width = Me.rtfPrint.Width
        Me.dgvAT.Height = Me.rtfPrint.Height
        Me.dgvAT.Refresh()

        Call ResizePan1a()
        Call ResizePan2a()

    End Sub

    Sub ResizePan1a()

        Dim tb As System.Windows.Forms.TextBox
        Dim ctl As System.Windows.Forms.Control

        Dim w1, w2

        w1 = Me.pan1a.Width

        Try
            For Each ctl In Me.pan1a.Controls

                If InStr(1, ctl.Name, "textbox", CompareMethod.Text) > 0 Then
                    Select Case ctl.Name
                        Case "CHARLINK1TextBox", "CHARLINK2TextBox"
                        Case Else
                            Try
                                w2 = ctl.Left
                                ctl.Width = w1 - w2 - 10
                            Catch ex As Exception

                            End Try
                    End Select

                End If

            Next
        Catch ex As Exception

        End Try

    End Sub

    Sub ResizePan2a()

        Dim tb As System.Windows.Forms.TextBox
        Dim ctl As System.Windows.Forms.Control

        Dim w1, w2

        w1 = Me.pan2a.Width

        Try
            For Each ctl In Me.pan2a.Controls

                If InStr(1, ctl.Name, "textbox", CompareMethod.Text) > 0 Then
                    Try
                        w2 = ctl.Left
                        ctl.Width = w1 - w2 - 10
                    Catch ex As Exception

                    End Try
                End If

            Next


        Catch ex As Exception

        End Try

    End Sub

    Private Sub frmAuditTrail_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize

        Call frmResize()

    End Sub

    Private Sub rbLong_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbLong.CheckedChanged

        Call DoView()

    End Sub

    Private Sub rbShort_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbShort.CheckedChanged

        Call DoView()

    End Sub

    Private Sub rbPrint_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbPrint.CheckedChanged

        Call DoView()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim str1 As String
        Dim int1 As Int32

        str1 = Me.rtfPrint.TextLength

        int1 = Me.rtfPrint.Lines.Length

        MsgBox(int1)

    End Sub

    Private Sub TBLAUDITTRAILBindingNavigator_RefreshItems(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBLAUDITTRAILBindingNavigator.RefreshItems

        Dim var1
        Dim strDt As String
        Dim str2 As String
        Dim dvr As DataRowView = Me.TBLAUDITTRAILBindingSource.Current 'this is a dataviewrow

        Try
            var1 = dvr.Item("DTSAVEDATE")
            str2 = "MMM dd, yyyy  hh:mm:ss tt"

            Try
                strDt = Format(CDate(var1), str2)
                Me.tDTSAVEDATETextBox.Text = strDt
            Catch ex As Exception
                Me.tDTSAVEDATETextBox.Text = ""
            End Try

        Catch ex As Exception
            Me.tDTSAVEDATETextBox.Text = ""
        End Try


    End Sub


    Private Sub cmdClipboard_Click(sender As System.Object, e As System.EventArgs) Handles cmdClipboard.Click

        Dim dgv As DataGridView = Me.dgvAT

        Dim intRows As Int64
        Dim intR1 As Int64
        Dim intR2 As Int64
        Dim intR As Short
        Dim strM As String

        intRows = dgv.Rows.Count

        strM = "Do you wish to copy all rows to the clipboard (Yes)?"
        strM = strM & ChrW(10) & ChrW(10) & "Or"
        strM = strM & ChrW(10) & ChrW(10) & "Do you wish to copy only the selected content (No)?"
        intR = MsgBox(strM, vbYesNoCancel, "Choose...")
        If intR = 6 Then 'yes
            dgv.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
            dgv.SelectAll()
            Clipboard.SetDataObject(dgv.GetClipboardContent())
        ElseIf intR = 7 Then 'no
            dgv.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
            Clipboard.SetDataObject(dgv.GetClipboardContent())
        Else 'cancel
            GoTo end1
        End If

end1:

    End Sub

End Class