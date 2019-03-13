Option Compare Text

Public Class frmVersionDescr

    Public boolCancel As Boolean = True


    Private Sub frmVersionDescr_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call ControlDefaults(Me)

        Dim str1 As String

        str1 = "Enter a Description of this Version" & ChrW(10) & "(limited to 2000 characters)"

        Me.lbl1.Text = str1


    End Sub

    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click

        If DoVal() Then
        Else
            boolCancel = False
            Me.Visible = False
        End If


    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click

        Me.Visible = False

    End Sub

    Function DoVal() As Boolean 'true = cancel

        DoVal = True

        Dim var1
        Dim strM As String

        strM = "Description cannot be blank."

        var1 = Me.rtbD.Text

        If Len(var1) = 0 Then
            MsgBox(strM, vbInformation, "Invalid action...")
        Else
            DoVal = False
        End If

    End Function
End Class