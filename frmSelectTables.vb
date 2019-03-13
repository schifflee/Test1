Option Compare Text

Public Class frmSelectTables
    Public boolCancel As Boolean = False

    Private Sub frmSelectTables_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Call ControlDefaults(Me)

        Dim str1 As String
        str1 = "This feature checks or unchecks all the Analytes shown in the Report Table grid by column or row."
        str1 = str1 & ChrW(10) & ChrW(10) & "Choose 'Columns' (Analytes) or 'Rows' (Tables)."
        str1 = str1 & ChrW(10) & "Choose 'Select All' or 'De-Select All'."
        str1 = str1 & ChrW(10) & "Select one or more Analytes or Tables."
        str1 = str1 & ChrW(10) & "Click OK or Cancel."

        Me.lblTitle.Text = str1

        Call Dothis()

    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click

        boolCancel = False
        Me.Visible = False

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        boolCancel = True
        Me.Visible = False

    End Sub

    Sub Dothis()

        Dim Count1 As Short
        Dim Count2 As Short
        Dim dv As System.Data.DataView
        Dim int1 As Short
        Dim int2 As Short
        Dim tblAnalytes As System.Data.DataTable
        Dim str1 As String
        Dim drows() As DataRow
        Dim dgv As DataGridView
        Dim dtbl1 As System.Data.DataTable
        Dim strF As String
        Dim strS As String
        Dim fs As Single
        Dim fn As String

        'Dim cs As DataGridColumnStyle
        'Note: a tablestyle already exists

        dgv = frmH.dgvReportTableConfiguration
        dv = dgv.DataSource

        tblAnalytes = tblAnalRefStandards
        strF = "ID_TBLSTUDIES = " & id_tblStudies
        strS = "ID_TBLSTUDIES ASC"
        'drows = tblAnalytes.Select(strF, strS)

        strF = "IsIntStd = 'No'"
        strS = "INTORDER ASC" ' "OriginalAnalyteDescription ASC"
        drows = tblAnalytesHome.Select(strF, strS)

        int1 = drows.Length 'debug

        Me.lbxAnalytes.Items.Clear()

        If Me.rbColumns.Checked Then

            fs = 12
            fn = Me.lbxAnalytes.Font.Name
            Me.lbxAnalytes.Font = New System.Drawing.Font(fn, fs)

            'Me.lbxAnalytes.DefaultFont.Size = 12

            int1 = drows.Length
            For Count2 = 0 To int1 - 1
                str1 = drows(Count2).Item("AnalyteDescription")

                If dgv.Columns.Contains(str1) Then 'include
                    'load in lbx
                    Me.lbxAnalytes.Items.Add(str1)

                End If

            Next
        Else

            fs = 10
            fn = Me.lbxAnalytes.Font.Name
            Me.lbxAnalytes.Font = New System.Drawing.Font(fn, fs)

            int1 = dgv.Rows.Count
            For Count2 = 0 To int1 - 1
                str1 = dgv("CHARHEADINGTEXT", Count2).Value
                Me.lbxAnalytes.Items.Add(str1)
            Next
        End If

    End Sub

    Private Sub rbColumns_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbColumns.CheckedChanged

        Call Dothis()

    End Sub

End Class