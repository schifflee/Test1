Option Compare Text

Public Class frmEULA
    Private checkPrint As Integer

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Visible = False
    End Sub

    Private Sub PrintDocument1_BeginPrint(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintEventArgs) Handles PrintDocument1.BeginPrint
        checkPrint = 0
    End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        ' Print the content of the RichTextBox. Store the last character printed.
        checkPrint = Me.rtxEULA.Print(checkPrint, Me.rtxEULA.TextLength, e)

        ' Look for more pages
        If checkPrint < Me.rtxEULA.TextLength Then
            e.HasMorePages = True
        Else
            e.HasMorePages = False
        End If
    End Sub

    Sub FormLoad()
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim strAppTitle As String

        'strAppTitle = NZ(System.Windows.Forms.Application.ProductName, "StudyDoc")
        strAppTitle = NZ(My.Application.Info.Title, "StudyDoc")

        boolEval = False

        str3 = "Use the following information to get technical support for " & strAppTitle
        str3 = str3 & vbCr & vbCr

        str1 = "LABIntegrity Support:"
        str2 = ""
        str3 = str3 & vbCr & str1 & vbTab & str2

        str1 = "Contact:"
        str2 = "Technical Support Department"
        str3 = str3 & vbCr & str1 & vbTab & vbTab & vbTab & vbTab & str2

        str1 = "Support Information:"
        str2 = "larry.elvebak@labintegrity.com"
        str3 = str3 & vbCr & str1 & vbTab & vbTab & str2

        str1 = "Support Telephone:"
        str2 = "(770)-573-0169"
        str3 = str3 & vbCr & str1 & vbTab & vbTab & str2

        str1 = "Product Information:"
        str2 = "www.labintegrity.com"
        str3 = str3 & vbCr & str1 & vbTab & vbTab & str2

        Me.txtSupport.Text = str3

        str1 = Me.txtEULA.Text
        str2 = Replace(str1, "[NAME]", strAppTitle, 1, -1, vbTextCompare)
        str2 = Replace(str2, "StudyDoc", strAppTitle, 1, -1, vbTextCompare)
        Me.txtEULA.Text = str2

        str1 = Me.txtTrial.Text
        str2 = Replace(str1, "[NAME]", strAppTitle, 1, -1, vbTextCompare)
        str2 = Replace(str2, "StudyDoc", strAppTitle, 1, -1, vbTextCompare)
        Me.txtTrial.Text = str2

    End Sub

    Private Sub frmEULA_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


    End Sub

    Private Sub cmdPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrint.Click

        If PrintDialog1.ShowDialog() = DialogResult.OK Then
            PrintDocument1.Print()
        End If

    End Sub
End Class