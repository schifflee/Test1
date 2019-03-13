Option Compare Text


Option Strict Off
Option Explicit Off

Imports System.Data
Imports System.Linq.Expressions
Imports System.Linq
Imports System.Data.DataTableExtensions
Imports System.Data.DataRowExtensions


Module modNoOptionExplicit


    Sub wStudyClicks(strN As String)

        If boolFormLoad Then
            GoTo end1 ' Exit Sub
        End If

        'clear data
        Call ClearData()


        Dim intR As Short
        Dim strM As String

        Dim rbtn As RadioButton
        Dim rbtnName As String = String.Empty

        Dim dgv As DataGridView

        dgv = frmH.dgvwStudy

        Dim boolF As Boolean = boolFormLoad
        boolFormLoad = True

        Dim ctl As Control
        For Each ctl In frmH.gbStudyFilter.Controls
            If TypeOf ctl Is RadioButton Then
                rbtn = DirectCast(ctl, RadioButton)
                If rbtn.Checked Then
                    rbtnName = rbtn.Name
                    If StrComp(rbtnName, strN, CompareMethod.Text) = 0 Then

                        boolRefresh = True

                        Call ConfigStudyTable(True, False)
                        frmH.cbxStudy.DataSource = frmH.dgvwStudy.DataSource

                        'select index 0 of cbxStudy
                        If frmH.cbxStudy.Items.Count = 0 Then

                        Else
                            frmH.cbxStudy.SelectedIndex = 0
                        End If

                        'select first row of dgv
                        If dgv.RowCount = 0 Then

                        Else
                            dgv.Rows(0).Selected = True

                            Call frmH.dgvwStudySelCh()

                            'frmH.pb1.Visible = False
                            'frmH.pb2.Visible = False
                            'frmH.lblProgress.Visible = False

                            frmH.panProgress.Visible = False
                            frmH.panProgress.Refresh()

                            boolStudyFired = True
                        End If
                    End If

                    Exit For
                End If
            End If
        Next

        'clear selection
        frmH.dgvwStudy.ClearSelection()

        'clear data
        Call ClearData()

        '20190130 LEE:
        Call SetStudyCount()

        boolFormLoad = boolF

end1:

    End Sub

    Sub ConfigStudyTable(ByVal boolW As Boolean, boolEntire As Boolean)

        Dim dgv As DataGridView
        Dim dv As System.Data.DataView
        Dim Count1 As Int32
        Dim Count2 As Int32
        Dim Count3 As Int32
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim strSQL As String
        Dim boolRO As Boolean
        'Dim rs As New ADODB.Recordset
        Dim int1 As Short
        Dim int2 As Short
        Dim intCol As Short
        Dim var1, var2, var3, var4
        Dim strF As String
        Dim strF1 As String
        Dim strF2 As String
        Dim strS As String

        tblReports.RejectChanges()

        'debug
        var1 = tblReports.Rows.Count
        var1 = var1

        Dim qryIDR '() As System.Data.DataRow

        'LINQ
        GoTo skipLINQ

        Try
            If frmH.rbArchive.Checked Then
                'need to create tblwStudies
            Else
                If frmH.optStudyDocStudies.Checked Then
                    qryIDR = From IDS In tblStudies.AsEnumerable() Join IDR In tblReports.AsEnumerable() On IDS("ID_TBLSTUDIES") Equals IDR("ID_TBLSTUDIES") Select IDS("INT_WATSONSTUDYID")
                ElseIf frmH.optStudyDocOpen.Checked Then
                    qryIDR = From IDS In tblStudies.AsEnumerable() Join IDR In tblReports.AsEnumerable() On IDS("ID_TBLSTUDIES") Equals IDR("ID_TBLSTUDIES") Where IDR("DTREPORTFINALISSUEDATE") Is DBNull.Value Select IDS("INT_WATSONSTUDYID")
                    ' Dim qryR = From IDS In tblStudies.AsEnumerable() Join IDR In tblReports On IDS("ID_TBLSTUDIES") Equals IDR("ID_TBLSTUDIES") Select IDS
                ElseIf frmH.optStudyDocClosed.Checked Then
                    qryIDR = From IDS In tblStudies.AsEnumerable() Join IDR In tblReports.AsEnumerable() On IDS("ID_TBLSTUDIES") Equals IDR("ID_TBLSTUDIES") Where IDR("DTREPORTFINALISSUEDATE") IsNot DBNull.Value Select IDS("INT_WATSONSTUDYID")
                    'Dim qryR = From IDS In tblStudies.AsEnumerable() Join IDR In tblReports On IDS("ID_TBLSTUDIES") Equals IDR("ID_TBLSTUDIES") Select IDS
                End If
            End If

        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

        Dim tblLINQ As DataTable
        'Try
        '    tblLINQ = qryIDR.todatatable
        'Catch ex As Exception
        '    var1 = var1
        'End Try



skipLINQ:

        Dim tbl1 As New System.Data.DataTable
        Dim col2 As New DataColumn

        col2.DataType = System.Type.GetType("System.Int32")
        col2.ColumnName = "STUDYID"
        col2.Caption = ""
        tbl1.Columns.Add(col2)

        Dim intC As Int64

        Dim dv1 As System.Data.DataView

        If tblwSTUDY.Columns.Contains("CHARREPORTTYPE") Then
        Else
            Try
                Dim col1 As New DataColumn
                col1.ColumnName = "CHARREPORTTYPE"
                col1.Caption = "Study Type"
                col1.DataType = System.Type.GetType("System.String")
                tblwSTUDY.Columns.Add(col1)
            Catch ex As Exception
                var1 = var1
            End Try
        End If


        int1 = tblwSTUDY.Rows.Count 'debug

        '20190130 LEE:
        'Put this as subroutine because frmAssignedSamples needs it
        tbl2 = CreatedgvwStudiesDatasource(True)

        ''Start CreatedgvwStudiesDatasource
        'Try
        '    dv1 = New DataView(tblwSTUDY.Clone, "", "STUDYNAME", DataViewRowState.CurrentRows)
        '    var1 = dv1.Count 'debug, should be 0 because clone doesn't bring over data
        'Catch ex As Exception
        '    var1 = ex.Message
        '    var1 = var1
        'End Try

        'Dim tbl2 As System.Data.DataTable
        'Try
        '    tbl2 = dv1.ToTable
        'Catch ex As Exception
        '    var1 = ex.Message
        '    var1 = tbl2.Rows.Count
        'End Try

        'var1 = tbl2.Rows.Count

        'If frmH.rbArchive.Checked Then
        '    Try

        '        tbl2 = tblwSTUDY.Copy
        '        var1 = tbl2.Rows.Count

        '        'need to add studytype (CHARREPORTTYPE) here
        '        var1 = tblwSTUDY.Rows(0).Item("STUDYID")
        '        'get id_tblStudies from tblstudies
        '        strF = "INT_WATSONSTUDYID = " & var1
        '        Dim rows2() As DataRow = tblStudies.Select(strF)
        '        var3 = rows2(0).Item("ID_TBLSTUDIES")
        '        strF1 = "ID_TBLSTUDIES = " & var3
        '        Dim rows1() As DataRow = tblReports.Select(strF1, "", DataViewRowState.CurrentRows)

        '        var1 = rows1(0).Item("CHARREPORTTYPE")
        '        tbl2.Rows(0).BeginEdit()
        '        tbl2.Rows(0).Item("CHARREPORTTYPE") = rows1(0).Item("CHARREPORTTYPE")
        '        tbl2.Rows(0).EndEdit()

        '    Catch ex As Exception
        '        var1 = ex.Message
        '        var1 = var1
        '    End Try
        'Else
        '    Try

        '        'Alternate code in case LINQ is too slow or is buggy
        '        'int1 = 0 'debug

        '        If frmH.optStudyDocStudies.Checked Then
        '            strF = "ID_TBLSTUDIES > 0"
        '        ElseIf frmH.optStudyDocOpen.Checked Then
        '            strF = "DTREPORTFINALISSUEDATE IS NULL"
        '        Else
        '            strF = "DTREPORTFINALISSUEDATE IS NOT NULL"
        '        End If


        '        Dim rows1() As DataRow = tblReports.Select(strF, "", DataViewRowState.CurrentRows)
        '        For Count1 = 0 To rows1.Length - 1

        '            ''debug
        '            'For Count2 = 0 To tblReports.Columns.Count - 1
        '            '    str1 = tblReports.Columns(Count2).ColumnName
        '            '    var1 = rows1(0).Item(str1)
        '            '    var1 = var1
        '            'Next

        '            var1 = rows1(Count1).Item("ID_TBLSTUDIES")
        '            strF1 = "ID_TBLSTUDIES = " & var1
        '            Dim rows2() As DataRow = tblStudies.Select(strF1)
        '            For Count2 = 0 To rows2.Length - 1
        '                var2 = rows2(Count2).Item("INT_WATSONSTUDYID")
        '                strF2 = "STUDYID = " & var2
        '                Dim rows3() As DataRow = tblwSTUDY.Select(strF2)
        '                For Count3 = 0 To rows3.Length - 1
        '                    int1 = int1 + 1 'debug

        '                    Dim dr1 As DataRow = tbl2.NewRow

        '                    dr1.BeginEdit()
        '                    dr1("PROJECTIDTEXT") = rows3(0).Item("PROJECTIDTEXT")
        '                    dr1("STUDYNAME") = rows3(0).Item("STUDYNAME")
        '                    dr1("STUDYNUMBER") = rows3(0).Item("STUDYNUMBER")
        '                    dr1("SPECIES") = rows3(0).Item("SPECIES")
        '                    dr1("STUDYTITLE") = rows3(0).Item("STUDYTITLE")
        '                    dr1("PROJECTID") = rows3(0).Item("PROJECTID")
        '                    dr1("STUDYID") = rows3(0).Item("STUDYID")
        '                    dr1("SPECIESID") = rows3(0).Item("SPECIESID")

        '                    Try
        '                        dr1("CHARREPORTTYPE") = rows1(Count1).Item("CHARREPORTTYPE")
        '                    Catch ex As Exception
        '                        var1 = var1
        '                    End Try
        '                    dr1.EndEdit()

        '                    tbl2.Rows.Add(dr1)

        '                Next
        '            Next
        '        Next

        '20190124 LEE:
        'screw LINQ, impossible to retrieve query results

        'Try
        '    int1 = 0 'debug
        '    For Each d In qryIDR

        '        var1 = d(0)
        '        var2 = d(1)
        '        var3 = d(2)
        '        var4 = d(3)

        '        var1 = d 'this is INT_WATSONID
        '        strF = "STUDYID = " & var1
        '        Dim row() As DataRow
        '        row = tblwSTUDY.Select(strF)

        '        If row.Length > 0 Then

        '            int1 = int1 + 1 'debug

        '            Dim dr1 As DataRow = tbl2.NewRow 'tbl2 = tblwSTUDY.Copy

        '            dr1.BeginEdit()
        '            dr1("PROJECTIDTEXT") = row(0).Item("PROJECTIDTEXT")
        '            dr1("STUDYNAME") = row(0).Item("STUDYNAME")
        '            dr1("STUDYNUMBER") = row(0).Item("STUDYNUMBER")
        '            dr1("SPECIES") = row(0).Item("SPECIES")
        '            dr1("STUDYTITLE") = row(0).Item("STUDYTITLE")
        '            dr1("PROJECTID") = row(0).Item("PROJECTID")
        '            dr1("STUDYID") = row(0).Item("STUDYID")
        '            dr1("SPECIESID") = row(0).Item("SPECIESID")

        '            Try
        '                dr1("CHARREPORTTYPE") = d("CHARREPORTTYPE")
        '            Catch ex As Exception
        '                var1 = var1
        '            End Try


        '            dr1.EndEdit()

        '            tbl2.Rows.Add(dr1)
        '        End If

        '    Next

        'Catch ex As Exception
        '    var1 = ex.Message
        '    var1 = var1
        'End Try


        '    Catch ex As Exception
        '    var1 = ex.Message
        '    var1 = var1
        'End Try
        'End If


        ''need to sort tbl2
        'Dim dv3 As DataView
        'Try
        '    dv3 = New DataView(tbl2, "", "STUDYNAME", DataViewRowState.CurrentRows)
        '    tbl2 = dv3.ToTable
        'Catch ex As Exception

        'End Try
        ''End CreatedgvwStudiesDatasource

        'Try
        '    dv1 = New DataView(tbl2, "", "STUDYNAME", DataViewRowState.CurrentRows)
        '    var1 = dv1.Count 'debug, should be 0 because clone doesn't bring over data
        'Catch ex As Exception
        '    var1 = ex.Message
        '    var1 = var1
        'End Try

        var1 = var1 'debug

        dgv = frmH.dgvwStudy

        Try
            Dim boolF As Boolean = boolFormLoad
            boolFormLoad = True
            dgv.DataSource = tbl2 ' dv
            boolFormLoad = boolF
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

        If boolEntire Then
        Else
            GoTo end1
        End If

        frmH.cbxStudy.DataSource = dgv.DataSource

        dgv.AllowUserToAddRows = False
        dgv.AllowUserToDeleteRows = False

        'hide all columns
        For Count1 = 0 To dgv.Columns.Count - 1
            dgv.Columns.Item(Count1).Visible = False
            dgv.Columns.Item(Count1).DisplayIndex = dgv.Columns.Count - 1
            dgv.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        'enter column headertexts
        Try
            For Count1 = 1 To dgv.ColumnCount - 1
                Dim dgc1 As New DataGridTextBoxColumn
                boolRO = True
                Select Case Count1
                    Case 1
                        str1 = "StudyName"
                        str2 = "Study Name"
                        boolRO = True
                    Case 2
                        str1 = "PROJECTIDTEXT"
                        str2 = "Project ID"
                        boolRO = True
                        intCol = Count1
                    Case 3
                        str1 = "StudyNumber"
                        str2 = "Study #"
                        boolRO = False 'True
                        'Case 4
                        '    str1 = "Species"
                        '    str2 = "Species"
                        '    boolRO = False
                    Case 4
                        str1 = "CHARREPORTTYPE"
                        str2 = "Study Type"
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
                dgv.Columns.Item(str1).Visible = boolRO ' True
                dgv.Columns.Item(str1).HeaderText = str2
                dgv.Columns.Item(str1).DisplayIndex = Count1 - 1
            Next
        Catch ex As Exception
            var1 = var1
        End Try


        'make studytitle column fit to grid
        If boolW Then
            dgv.Columns.Item("StudyTitle").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            dgv.Columns.Item("StudyTitle").DefaultCellStyle.WrapMode = DataGridViewTriState.True
            dgv.AutoResizeColumns()
            dgv.AutoResizeRows()
        End If

        'set first row as current row
        'NO!!
        'dgv.CurrentCell = dgv.Rows.Item(0).Cells(intCol)

        'assign this datasource to cbxstudy
        frmH.cbxStudy.DataSource = dgv.DataSource 'dv
        frmH.cbxStudy.DisplayMember = "STUDYNAME"
        'cbxStudy.ValueMember = "ID_TBLSTUDIES"
        'set selection to nothing 
        Try
            If frmH.rbArchive.Checked Then
                frmH.cbxStudy.SelectedIndex = 0

                'this doesn't seem to be triggering.
                'set dgv instead
                dgv.Rows(0).Cells(0).Selected = True
                int1 = dgv.RowCount
                If IsNothing(dgv.CurrentRow) Then

                Else
                    int1 = dgv.CurrentRow.Index
                End If

                'Call frmH.dgvwStudySelCh()

                'frmH.pb1.Visible = False
                'frmH.pb2.Visible = False
                'frmH.lblProgress.Visible = False

                'frmH.panProgress.Visible = False
                'frmH.panProgress.Refresh()

                boolStudyFired = True
            Else
                frmH.cbxStudy.SelectedIndex = -1
            End If
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

        'now fill tblastudy with blank data
        tblASTUDY = tbl2.Copy ' tblwSTUDY.Copy
        'remove all rows from tblAStudy
        tblASTUDY.Clear()


end1:


        Call ResetStudyRecord()


        'initially have no rows chosen in dgv
        Dim boolFL As Boolean = boolFormLoad
        Try
            boolFormLoad = True
            dgv.ClearSelection()
        Catch ex As Exception

        End Try

        boolFormLoad = boolFL

        'clear txtfilter
        frmH.txtFilterStudy.Clear()

    End Sub

    Private Function cbxStudy() As Object
        Throw New NotImplementedException
    End Function

End Module
