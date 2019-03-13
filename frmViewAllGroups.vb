Option Compare Text

Public Class frmViewAllGroups

    Dim boolFormLoad As Boolean = False
    Dim boolHold As Boolean

    Private Sub frmViewAllGroups_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call DoubleBufferControl(Me, "dgv")

        boolFormLoad = True

        Call LoadTables()

        boolFormLoad = False

    End Sub

    Sub LoadTables()

        '  - tblAnalyteGroups
        '      a list of each analyte/matrix/calibration set
        '      this table is created in Sub EstablishCalStdGroups
        '          tblAnalyteGroups = dv1a.ToTable("a", True, "ANALYTEDESCRIPTION", "ANALYTEID", "INTSTD", "INTGROUP", "ANALYTEDESCRIPTION_C", "MATRIX", "INTCALSET", "CALIBRSET") + "INTORDER"

        '  - tblCalStdGroupsAll
        '      examples of each calibration standard calibration set associated with each group for ALL analytical runs
        '      see Sub CreatetblCalStdGroups for a list of columns in this table

        '  - tblCalStdGroupsAcc
        '      examples of each calibration standard calibration set associated with each group for ACCEPTED analytical runs
        '      same columns as tblCalStdGroupsAll

        '  - tblCalStdGroupAssayIDsAll
        '      a list of all the RunIDs (AssayID) associated with each group for ALL analytical runs
        '      see Sub CreatetblCalStdGroupAssayIDs for a list of columns in this table

        '  - tblCalStdGroupAssayIDsAcc
        '      a list of all the RunIDs (AssayID) associated with each group for ALL analytical runs
        '      same columns as tblCalStdGroupAssayIDsAll

        Dim lbx As ListBox = Me.lbxGroupTables
        Dim Count1 As Short
        Dim str1 As String

        boolHold = True

        lbx.Items.Clear()

        For Count1 = 1 To 5
            Select Case Count1
                Case 1
                    str1 = "tblAnalyteGroups"
                Case 2
                    str1 = "tblCalStdGroupsAll"
                Case 3
                    str1 = "tblCalStdGroupsAcc"
                Case 4
                    str1 = "tblCalStdGroupAssayIDsAll"
                Case 5
                    str1 = "tblCalStdGroupAssayIDsAcc"
            End Select

            lbx.Items.Add(str1)

        Next

        boolHold = False

        'select first itme in lbx
        lbx.SelectedIndex = 0

    End Sub

    Private Sub lbxGroupTables_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lbxGroupTables.SelectedIndexChanged

        If boolHold Then
            Exit Sub
        End If

        Call LoadTable()

    End Sub

    Sub LoadTable()

        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strS As String
        Dim lbx As ListBox = Me.lbxGroupTables
        Dim tbl As DataTable
        Dim dv As DataView
        Dim dgv As DataGridView = Me.dgvFormGroups

        str1 = lbx.SelectedItem

        str2 = " Table Content"
        str4 = ""
        Select Case str1
            Case "tblAnalyteGroups"
                tbl = tblAnalyteGroups
            Case "tblCalStdGroupsAll"
                tbl = tblCalStdGroupsAll
            Case "tblCalStdGroupsAcc"
                tbl = tblCalStdGroupsAcc
            Case "tblCalStdGroupAssayIDsAll"
                tbl = tblCalStdGroupAssayIDsAll
            Case "tblCalStdGroupAssayIDsAcc"
                tbl = tblCalStdGroupAssayIDsAcc
        End Select

        str4 = ": " & tbl.Rows.Count & " rows"
        str3 = str1 & str2 & str4
        Me.lblGroupTableContent.Text = str3



        Dim var1
        Try
            dv = New DataView(tbl)
            strS = ReturnSort(True)
            dv.Sort = strS
            dv.AllowDelete = False
            dv.AllowEdit = False
            dv.AllowNew = False

            dgv.DataSource = dv

            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgv.AutoResizeColumns()
        Catch ex As Exception
            var1 = ex.Message
        End Try

    End Sub

    Private Sub cmdCopyAll_Click(sender As Object, e As EventArgs) Handles cmdCopyAll.Click

        Dim var1, var2
        Dim dgv As DataGridView = Me.dgvFormGroups
        Dim lbx As ListBox = Me.lbxGroupTables
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim int1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String

        int1 = lbx.Items.Count
        var2 = ChrW(10)
        For Count1 = 0 To int1 - 1

            lbx.SelectedIndex = Count1
            str1 = lbx.SelectedItem
            var1 = str1

            'do column headers
            str2 = ""
            For Count2 = 0 To dgv.ColumnCount - 1
                str3 = dgv.Columns(Count2).Name
                If Count2 = 0 Then
                    str2 = str3
                Else
                    str2 = str2 & ChrW(9) & str3
                End If
            Next
            var1 = var1 & ChrW(10) & str2

            'do contents
            For Count3 = 0 To dgv.Rows.Count - 1
                str2 = ""
                For Count2 = 0 To dgv.ColumnCount - 1
                    str3 = NZ(dgv(Count2, Count3).Value, "")
                    If Count2 = 0 Then
                        str2 = str3
                    Else
                        str2 = str2 & ChrW(9) & str3
                    End If
                Next
                var1 = var1 & ChrW(10) & str2
            Next

            var2 = var2 & ChrW(10) & var1 & ChrW(10)


        Next

        'place var2 in clipboard
        Try
            Clipboard.Clear()
        Catch ex As Exception

        End Try
        Try
            Clipboard.SetText(var2.ToString)
        Catch ex As Exception

        End Try


        MsgBox("Information pasted to clipboard.", vbInformation, "Information pasted to clipboard...")

    End Sub
End Class