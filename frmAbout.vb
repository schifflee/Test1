Option Compare Text

Public Class frmAbout

    Private Sub frmAbout_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Call ControlDefaults(Me)

        Dim str1 As String
        Dim var1, var2, var3, var4, var5

        Me.lblR.Text = ChrW(174)
        'str1 = "About Gubbs GuWu" & ChrW(174) & "..."
        str1 = "About LABIntegrity StudyDoc" & ChrW(8482) & "..."
        Me.Text = str1

        str1 = GetStudyDocHeader(True)
        Me.lbl2.Text = str1

        'find last decimal
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim Count1 As Short
        'var1 = "Version " & System.Windows.Forms.Application.ProductVersion
        'int1 = 1
        'Do Until InStr(int1, var1, ".", CompareMethod.Text) = 0
        '    int2 = InStr(int1, var1, ".", CompareMethod.Text)
        '    If int2 = 0 Then
        '        Exit Do
        '    End If
        '    int1 = int2 + 1
        'Loop
        'var1 = Mid(var1, 1, int1 - 2)

        str1 = NZ(System.Windows.Forms.Application.ProductName, "StudyDoc")
        Me.lblVS.Text = str1

        var1 = GetVersionFour()
        var2 = My.Application.Info.Description

        str1 = var1 & "    " & var2
        Me.lbl3.Text = str1


        str1 = "Manages study design projects and generates study reports."
        str1 = "Manages and generates Watson study reports."
        Me.lbl4.Text = str1

        str1 = "Warning: This computer program is protected by copyright law and international treaties."
        str1 = str1 & " Unauthorized reproduction or distribution of this program, or any portion of it, may"
        str1 = str1 & " result in severe civil and criminal penalties, and will be prosecuted to the "
        str1 = str1 & "maximum extent possible under the law."
        Me.lbl5.Text = str1

        var1 = Format(Now(), "yyyy")
        str1 = "Copyright " & Chr(169) & "  2014-" & var1 & " LABIntegrity"
        Me.lbl6.Text = str1

        Dim l, r, t, b, w, h

        'position form
        w = My.Computer.Screen.WorkingArea.Width
        h = My.Computer.Screen.WorkingArea.Height

        'Me.Left = w * 0.1
        'Me.Width = w * 0.8
        'Me.Top = h * 0.1
        'Me.Height = h * 0.8

        'l = Me.wb1.Left
        'w = Me.Width
        'Me.wb1.Width = w - l - l
        'h = Me.Height
        't = Me.wb1.Top
        'Me.wb1.Height = h - t - 100

        str1 = "www.labintegrity.com" ' "www.gubbsinc.com"
        'Try
        '    Me.wb1.Navigate(str1)
        'Catch ex As Exception

        'End Try


    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Visible = False
    End Sub

    Private Sub cmdBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBack.Click
        Me.wb1.GoBack()
    End Sub

    Private Sub Forward_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Forward.Click
        Me.wb1.GoForward()
    End Sub

    Private Sub cmdRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRefresh.Click
        Me.wb1.Refresh()
    End Sub

    Private Sub cmdLicense_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLicense.Click

        Dim str1 As String
        Dim frm As New frmEULA

        Call frm.FormLoad()

        If boolEval Then
            frm.rtxEULA.Text = frm.txtTrial.Text
            str1 = "Software Trial Version End User License Agreement (EULA)"
        Else
            frm.rtxEULA.Text = frm.txtEULA.Text
            str1 = "Software End User License Agreement (EULA)"
        End If
        frm.Text = str1
        frm.ShowDialog()

        frm.Dispose()

        Me.wb1.Visible = True

    End Sub

    Private Sub cmdSupport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSupport.Click

        Dim str1 As String
        Dim frm As New frmEULA

        Call frm.FormLoad()

        'frm.rtxEULA.Font.Size = 12

        'frm.rtxEULA.Font.Bold = True

        frm.rtxEULA.Text = frm.txtSupport.Text
        str1 = "Support Information"
        frm.Text = str1
        frm.ShowDialog()

        frm.Dispose()

        Me.wb1.Visible = True

    End Sub

    Private Sub lblGubbsInc_Click(sender As Object, e As EventArgs) Handles lblGubbsInc.Click

    End Sub

    Private Sub lbl5_Click(sender As Object, e As EventArgs) Handles lbl5.Click

    End Sub

    Private Sub lbl3_Click(sender As Object, e As EventArgs) Handles lbl3.Click

    End Sub

    Private Sub lblWWW_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lblWWW.LinkClicked

        Try
            VisitLink()
        Catch ex As Exception
            ' The error message
            MessageBox.Show("Unable to open link that was clicked.")
        End Try


    End Sub

    Sub VisitLink()
        ' Change the color of the link text by setting LinkVisited 
        ' to True.
        Me.lblWWW.LinkVisited = True
        ' Call the Process.Start method to open the default browser 
        ' with a URL:
        System.Diagnostics.Process.Start("http://www.labintegrity.com")
    End Sub

End Class