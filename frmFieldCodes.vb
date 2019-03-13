Option Compare Text

Imports System.Windows.Forms

Public Class frmFieldCodes
    Public boolCancel As Boolean = True
    Public strFC As String = ""
    Public posX As Double
    Public posY As Double
    Public boolFormLoad As Boolean = True
    Public boolCopyAll As Boolean = False
    Public boolFromShowFC As Boolean = False

    Sub CountFC()

        Dim dgv As DataGridView = Me.dgvFC
        Dim intRows As Int64 = dgv.Rows.Count
        Dim var1


        Try
            Dim str1 = Format(intRows, "#,###")
            Me.txtCount.Text = str1
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try


    End Sub

    Private Sub frmFieldCodes_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim dtbl As System.Data.DataTable
        Dim dgv As DataGridView
        Dim intRows As Short
        Dim intCols As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim str1 As String
        Dim strS As String
        Dim strF As String

        Call DoubleBufferControl(Me, "dgv")
        Call ControlDefaults(Me)

        'str1 = "NOTE:  Some field codes are missing." & ChrW(10) & "Field codes relating to report templates (e.g. table insertions, table references) are not listed."
        'str1 = str1 & ChrW(10) & ChrW(10) & "Open a study if you wish to display the full list of field codes."

        str1 = "NOTE:  Some field codes are missing." & ChrW(10) & "Study-specific field codes (e.g. table insertions, table references) are not listed."
        str1 = str1 & ChrW(10) & ChrW(10) & "Open a study if you wish to display the full list of field codes."

        Me.lblReportTemplateFCstatus.Text = str1

        Cursor.Current = Cursors.WaitCursor

        boolFormLoad = True

        'Call AddSigBlocks()
        'Call AddReportTemplates()
        ''Call AddTables()
        'Call AddAppFigs()

        dtbl = tblFieldCodes
        dgv = Me.dgvFC

        ''position and size form and dgv
        'Me.Cursor = New Cursor(Cursor.Current.Handle)
        'Me.Left = 0
        'Me.Top = Cursor.Position.Y

        Dim w, h, w1, h1, l
        w = Me.Width
        h = Me.Height
        w = My.Computer.Screen.WorkingArea.Width
        h = My.Computer.Screen.WorkingArea.Height
        Me.Top = 0
        Me.Left = 0
        Me.Width = w
        Me.Height = h

        'Me.panFC.Width = w - (Me.panFC.Left + (w - Me.TableLayoutPanel1.Left - Me.TableLayoutPanel1.Width))
        'w1 = Me.panFC.Width

        'Me.panFC.Height = Me.TableLayoutPanel1.Top - Me.panFC.Top - 10
        Me.panFC.Height = h - Me.panFC.Top - 100
        h1 = Me.panFC.Height
        l = Me.panFC.Left
        Me.panFC.Width = w - (Me.panFC.Left * 2)

        'Me.Cancel_Button.Left = Me.panFC.Left + Me.panFC.Width - Me.Cancel_Button.Width
        'Me.OK_Button.Left = Me.Cancel_Button.Left - Me.OK_Button.Width - 6

        'Me.OK_Button.Top = 12
        'Me.Cancel_Button.Top = 12

        Me.OK_Button.Left = Me.lblGroup.Left
        Me.OK_Button.Top = Me.lblGroup.Top + Me.lblGroup.Height + 15

        Me.Cancel_Button.Top = Me.OK_Button.Top
        Me.Cancel_Button.Left = Me.OK_Button.Left + Me.OK_Button.Width + 5

        'Me.gbCopyAll.Top = Me.Cancel_Button.Top + Me.cmdCopyAll.Top
        Me.gbCopyAll.Left = Me.Cancel_Button.Left + Me.Cancel_Button.Width + 5

        Me.gbxlblReportTemplateFCstatus.Top = cbxGroup.Top - 3
        'MsgBox(Cursor.Position.X & ", " & Cursor.Position.Y)

        dgv.AllowUserToOrderColumns = True
        dgv.AllowUserToResizeColumns = True
        dgv.AllowUserToResizeRows = True

        'dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        'dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        'dgv.SelectionMode = DataGridViewSelectionMode.CellSelect

        '******

        dgv.AllowUserToResizeColumns = True
        dgv.AllowUserToResizeRows = True
        dgv.RowHeadersWidth = 25
        'dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        'dgv.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

        '******

        strF = "CHARFIELDCODE NOT LIKE '*role*'"
        strS = "CHARFIELDCODE ASC"
        Dim dv As system.data.dataview = New DataView(dtbl, strF, strS, DataViewRowState.CurrentRows)

        'configure dgv
        dv.AllowDelete = False
        dv.AllowEdit = False
        dv.AllowNew = False
        strS = "CHARFIELDCODE ASC"
        dv.Sort = strS

        dgv.DataSource = dv
        dgv.ReadOnly = True

        intRows = dgv.RowCount
        intCols = dgv.ColumnCount

        'first hide all columns
        For Count1 = 0 To intCols - 1
            dgv.Columns(Count1).Visible = False
            dgv.Columns(Count1).SortMode = DataGridViewColumnSortMode.Automatic
        Next

        'ID_TBLFIELDCODES
        'CHARFIELDCODE
        'CHARDESCRIPTION
        'CHARTAB
        'CHARTABLE
        'CHAREXAMPLE
        'CHARGROUP
        'UPSIZE_TS

        str1 = "CHARFIELDCODE"
        dgv.Columns(str1).Visible = True
        dgv.Columns(str1).HeaderText = "Field Code"

        str1 = "CHARDESCRIPTION"
        dgv.Columns(str1).Visible = True
        dgv.Columns(str1).HeaderText = "Description"

        str1 = "CHARTAB"
        dgv.Columns(str1).Visible = True
        dgv.Columns(str1).HeaderText = "Tab"

        str1 = "CHARTABLE"
        dgv.Columns(str1).Visible = True
        dgv.Columns(str1).HeaderText = "Table"

        str1 = "CHAREXAMPLE"
        dgv.Columns(str1).Visible = True
        dgv.Columns(str1).HeaderText = "Example"
        dgv.Columns(str1).DisplayIndex = 3

        str1 = "CHARGROUP"
        dgv.Columns(str1).Visible = True
        dgv.Columns(str1).HeaderText = "Group"

        Call SizeColumns(dgv)

        'populate cbxGroup
        Call FillcbxGroup()

        If (BOOLTEMPLATEFIELDCODESLOADED) Then
            gbxlblReportTemplateFCstatus.Visible = False
        Else
            gbxlblReportTemplateFCstatus.Visible = True
        End If

        boolFormLoad = False

        Call CountFC()

        Cursor.Current = Cursors.Default

    End Sub

    Sub FillcbxGroup()

        Dim dgv As DataGridView
        Dim dv As system.data.dataview

        dgv = Me.dgvFC
        dv = dgv.DataSource

        Dim tblG As System.Data.DataTable = dv.ToTable("a", True, "CHARGROUP")
        Dim dv1 As system.data.dataview
        Dim int1 As Short
        Dim var1
        Dim intTRows As Short
        Dim tblR As System.Data.DataTable
        Dim rowsR() As DataRow
        Dim strA As String
        Dim strB As String
        Dim Count1 As Short
        Dim Count2 As Short
        Dim strF As String
        Dim str1 As String


        tblR = tblTemplates
        strF = "BOOLACTIVE = -1"
        rowsR = tblR.Select(strF)
        intTRows = rowsR.Length

        'dv1 = tblG.DefaultView
        dv1 = New DataView(tblG)
        dv1.Sort = "CHARGROUP ASC"
        int1 = tblG.Rows.Count
        Me.cbxGroup.Items.Clear()
        Me.cbxGroup.Items.Add("[NONE]")
        Me.cbxGroup.Items.Add("--------------------------------------------")

        'For Count1 = 0 To intTRows - 1
        '    str1 = "Report Template - " & rowsR(Count1).Item("CHARTEMPLATENAME")
        '    Me.cbxGroup.Items.Add(str1)
        'Next

        'Me.cbxGroup.Items.Add("--------------------------------------------")

        'For Count2 = 1 To 4

        '    Select Case Count2
        '        Case 1
        '            strA = "RepTemplate_"
        '            strB = "Tables for Report Template - "
        '        Case 2
        '            strA = "RunID_"
        '            strB = "Run IDs of runs for Report Template - "
        '        Case 3
        '            strA = "FirstAnalysisDate_"
        '            strB = "First Analysis Date of runs for Report Template - "
        '        Case 4
        '            strA = "LastAnalysisDate_"
        '            strB = "Last Analysis Date of runs for Report Template - "
        '    End Select

        '    For Count1 = 0 To intTRows - 1
        '        If Count1 = 0 Then
        '            str1 = "Current Report"
        '            strA = strB & str1
        '            Me.cbxGroup.Items.Add(strA)
        '        End If
        '        str1 = rowsR(Count1).Item("CHARTEMPLATENAME")
        '        strA = strB & str1
        '        Me.cbxGroup.Items.Add(strA)
        '    Next

        '    Me.cbxGroup.Items.Add("--------------------------------------------")

        'Next

        Me.cbxGroup.Items.Add("Signature Block")

        Me.cbxGroup.Items.Add("--------------------------------------------")

        Me.cbxGroup.Items.Add("Tables - Individual")

        Me.cbxGroup.Items.Add("--------------------------------------------")

        For Count1 = 0 To int1 - 1
            var1 = dv1(Count1).Item("CHARGROUP")
            If IsDBNull(var1) Then
            ElseIf Len(var1) = 0 Then
            Else
                If InStr(1, var1, "for Report Template", CompareMethod.Text) > 0 Then
                Else
                    Me.cbxGroup.Items.Add(var1)
                End If
            End If
        Next

        Me.cbxGroup.SelectedIndex = 0

    End Sub

    Sub SizeColumns(ByVal dgv As DataGridView)

        dgv.AutoResizeColumns()
        'bring back chardescription

        Dim var1, w

        var1 = dgv.Columns("CHARFIELDCODE").Width
        If var1 > dgv.Width * 0.2 Then
            'dgv.Columns("CHARFIELDCODE").Width = dgv.Width * 0.2
        End If
        w = var1

        var1 = dgv.Columns("CHARDESCRIPTION").Width
        If var1 > (dgv.Width - w) * 0.5 Then
            dgv.Columns("CHARDESCRIPTION").Width = (dgv.Width - w) * 0.5
        End If

        var1 = dgv.Columns("CHAREXAMPLE").Width
        If var1 > (dgv.Width - w) * 0.4 Then
            dgv.Columns("CHAREXAMPLE").Width = (dgv.Width - w) * 0.4
        End If

        dgv.RowHeadersWidth = 25

        dgv.AutoResizeRows()



    End Sub


    Sub OK()
        If Me.dgvFC.Rows.Count = 0 Then
            MsgBox("An item must be chosen.", MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If
        If Me.dgvFC.CurrentRow Is Nothing Then
            MsgBox("An item must be chosen.", MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If

        Me.DialogResult = System.Windows.Forms.DialogResult.OK

        Dim intRow As Short
        If Me.dgvFC.Rows.Count = 0 Then
        Else
            intRow = Me.dgvFC.CurrentRow.Index
            strFC = Me.dgvFC("CHARFIELDCODE", intRow).Value
        End If

        boolCancel = False
        Me.Close()

    End Sub


    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Call OK()

    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.strFC = ""

        ''reject changes in tblFieldCode
        'tblFieldCodes.RejectChanges()

        boolCancel = True
        Me.Close()
    End Sub

    Private Sub cbxGroup_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxGroup.SelectedIndexChanged

        Call FilterStuff()

    End Sub

    Private Sub txtFilterFC_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFilterFC.TextChanged
        Call FilterStuff()
        Me.txtFilterFC.Focus()
    End Sub

    Sub FilterStuff()

        If boolFormLoad Then
            Exit Sub
        End If

        Dim strGroup As String
        Dim strFC As String
        Dim strDescr As String
        Dim strTable As String
        Dim boolG As Boolean = False
        Dim boolF As Boolean = False
        Dim boolD As Boolean = False
        Dim boolT As Boolean = False
        Dim boolNone As Boolean = True
        Dim boolSS As Boolean = False

        Cursor.Current = Cursors.WaitCursor

        If Me.rbNone.Checked Then
            boolNone = True
        Else
            boolNone = False
        End If
        If Me.rbStudySpecific.Checked Then
            boolSS = True
        Else
            boolSS = False
        End If

        strGroup = Me.cbxGroup.SelectedItem

        strFC = Me.txtFilterFC.Text
        If Len(strFC) = 0 Then
        Else
            strFC = "*" & Me.txtFilterFC.Text & "*"
        End If

        strDescr = Me.txtFilterDescr.Text
        If Len(strDescr) = 0 Then
        Else
            strDescr = "*" & Me.txtFilterDescr.Text & "*"
        End If

        strTable = Me.txtFilterTable.Text
        If Len(strTable) = 0 Then
        Else
            strTable = "*" & Me.txtFilterTable.Text & "*"
        End If

        If StrComp(strGroup, "[NONE]", CompareMethod.Text) = 0 Then
        Else
            boolG = True
        End If

        'account for asterisks
        If Len(strFC) < 4 Then
        Else
            boolF = True
        End If

        If Len(strDescr) < 4 Then
        Else
            boolD = True
        End If

        If Len(strTable) < 4 Then
        Else
            boolT = True
        End If

        'boolF = True
        'boolD = True
        'boolT = True

        Dim dv As system.data.dataview
        Dim dgv As DataGridView
        Dim strF As String
        Dim strS As String
        Dim str1 As String

        dgv = Me.dgvFC
        dv = dgv.DataSource

        strS = "CHARFIELDCODE ASC"

        strF = ""

        If boolG Then
        Else
            'If Not (boolF) And Not (boolD) And Not (boolT) Then
            '    strS = "ID_TBLFIELDCODES ASC"
            '    strF = ""
            '    dv.RowFilter = "ID_TBLFIELDCODES > -1"
            '    dv.Sort = strS
            '    GoTo end2
            'End If
        End If

        If boolG And boolF And boolD And boolT Then 'exit
            GoTo end1
        End If

        strF = ""
        If boolSS Then
            strF = "CHARGROUP = 'Table-Specific variables'"
        Else
            If boolNone Then
            Else
                strF = "CHARGROUP <> 'Table-Specific variables'"
            End If
        End If

        If boolG Then
            strF = "CHARGROUP = '" & strGroup & "'"
        End If

        If boolF Then
            If Len(strF) = 0 Then
                strF = "CHARFIELDCODE LIKE '" & strFC & "'"
            Else
                strF = strF & " AND CHARFIELDCODE LIKE '" & strFC & "'"
            End If
        End If

        If boolD Then
            If Len(strF) = 0 Then
                strF = "CHARDESCRIPTION LIKE '" & strDescr & "'"
            Else
                strF = strF & " AND CHARDESCRIPTION LIKE '" & strDescr & "'"
            End If
        End If

        If boolT Then
            'If Len(strF) = 0 Then
            '    strF = "CHARTABLE LIKE '" & strTable & "'"
            'Else
            '    strF = strF & " AND CHARTABLE LIKE '" & strTable & "'"
            'End If
            '20181128 LEE:
            'txtFilterTable re-purposed to be additional Field Code
            If Len(strF) = 0 Then
                strF = "CHARFIELDCODE LIKE '" & strTable & "'"
            Else
                strF = strF & " AND CHARFIELDCODE LIKE '" & strTable & "'"
            End If
        End If

        If StrComp(strGroup, "Tables - Individual", CompareMethod.Text) = 0 Then
            strS = "ID_TBLFIELDCODES ASC"
        End If

        'correct for special characters
        If InStr(1, strF, "[", CompareMethod.Text) Then
            strF = Replace(strF, "[", "[[]", 1, -1, CompareMethod.Text)
        ElseIf InStr(1, strF, "]", CompareMethod.Text) Then
            strF = Replace(strF, "]", "[]]", 1, -1, CompareMethod.Text)
        End If


        Try
            dv.RowFilter = strF
            dv.Sort = strS

            'Call SizeColumns(dgv)

        Catch ex As Exception
            str1 = "The filter string:" & ChrW(10) & ChrW(10)
            str1 = str1 & strF & ChrW(10) & ChrW(10)
            str1 = str1 & "is invalid"
            MsgBox(str1, MsgBoxStyle.Information, "Invalid action")

        End Try

end1:

        dgv.AutoResizeRows()
        'dgv.AutoResizeColumns()

end2:

        dgv.Select()

        Call CountFC()

        Cursor.Current = Cursors.Default

    End Sub

    Private Sub txtFilterDescr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFilterDescr.TextChanged
        Call FilterStuff()
        Me.txtFilterDescr.Focus()
    End Sub

    Private Sub txtFilterTable_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFilterTable.TextChanged
        Call FilterStuff()
        Me.txtFilterTable.Focus()
    End Sub


    Private Sub dgvFC_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvFC.CellDoubleClick

        Call OK()

    End Sub

    Private Sub dgvFC_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvFC.Sorted

        Me.dgvFC.AutoResizeRows()

    End Sub

    Private Sub rbNone_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbNone.CheckedChanged
        If Me.rbNone.Checked Then
            Call FilterStuff()
        End If
    End Sub

    Private Sub rbReportItems_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbReportItems.CheckedChanged
        If Me.rbReportItems.Checked Then
            Call FilterStuff()
        End If
    End Sub

    Private Sub rbStudySpecific_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbStudySpecific.CheckedChanged
        If Me.rbStudySpecific.Checked Then
            Call FilterStuff()
        End If
    End Sub

    Private Sub cmdCopyAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCopyAll.Click

        Call CopyAll()

    End Sub

    Sub CopyAll()

        Dim dgv As DataGridView
        Dim Count1 As Int16
        Dim intRows As Int16
        Dim intRow As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim int1 As Int16 = 0
        Dim arrFC(100) As String
        Dim intFC As Short = 0
        Dim strM As String

        Dim boolWL As Boolean = Me.rbWithLabels.Checked

        dgv = Me.dgvFC
        intRows = dgv.RowCount

        If intRows = 0 Then
            MsgBox("Field Codes must be shown", MsgBoxStyle.Information, "Invalid action...")
            GoTo end1
        End If

        Dim strT As String = ""

        strFC = ""

        For Count1 = 0 To intRows - 1
            str1 = Me.dgvFC("CHARFIELDCODE", Count1).Value

            If boolFromShowFC Then
                str4 = Me.dgvFC("CHARDESCRIPTION", Count1).Value
                str4 = Replace(str4, ChrW(9), " ", 1, -1, CompareMethod.Text)
                str4 = Replace(str4, ChrW(13), ChrW(11), 1, -1, CompareMethod.Text)
                str4 = Replace(str4, ChrW(10), ChrW(11), 1, -1, CompareMethod.Text)
                str4 = Replace(str4, "'", """", 1, -1, CompareMethod.Text)
                str4 = Replace(str4, ChrW(19), "", 1, -1, CompareMethod.Text)

                str2 = Mid(str1, 2, Len(str1) - 2) & ":  " & ChrW(9)
            Else
                str2 = Mid(str1, 2, Len(str1) - 2) & ":  "
            End If


            If InStr(1, str1, "TableINSERTTable_", CompareMethod.Text) > 0 Then
                intFC = intFC + 1
                arrFC(intFC) = str1
            Else
                If boolWL Then
                    If Me.boolFromShowFC Then
                        str3 = str2 & str1 & ChrW(9) & str4
                    Else
                        str3 = str2 & str1
                    End If

                Else
                    str3 = str1
                End If
                int1 = int1 + 1
                If int1 = 1 Then
                    strFC = str3 & ChrW(10)
                Else
                    strFC = strFC & str3 & ChrW(10)
                End If
            End If
        Next

        If intFC = 0 Then
        Else
            For Count1 = 1 To intFC
                strT = arrFC(Count1)
                strFC = strFC & ChrW(10) & strT & ChrW(10)
            Next

        End If

        If boolFromShowFC Then
            'place strfc in clipboard
            Try
                Clipboard.Clear()
            Catch ex As Exception

            End Try
            Try
                Clipboard.SetText(strFC)
                strM = "The contents of the table have been placed in the clipboard."
            Catch ex As Exception
                strM = "There was a problem placing the contents of the table in the clipboard."
            End Try


            MsgBox(strM, vbInformation, "Copied...")
        Else
            boolCancel = False
            Me.Close()
        End If


end1:
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        MsgBox(Me.cmdCopyAll.Visible)

    End Sub
End Class
