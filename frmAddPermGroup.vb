Option Compare Text

Public Class frmAddPermGroup

    Public boolCancel As Boolean = True

    Private Sub frmAddPermGroup_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Call DoubleBufferControl(Me, "dgv")

        Call ControlDefaults(Me)

        Dim str1 As String

        str1 = "Choose an existing Permissions Group to use as a base setting:"
        str1 = "Choose an existing Permissions Group from which to apply initial settings:"

        Me.Location = New Point(100, 100)


    End Sub

    Private Sub cmdCancel_Click(sender As System.Object, e As System.EventArgs) Handles cmdCancel.Click

        boolCancel = True
        Me.Visible = False

    End Sub

    Private Sub cmdOK_Click(sender As System.Object, e As System.EventArgs) Handles cmdOK.Click

        If ValidateThis() Then

            boolCancel = False
            Me.Visible = False

        End If


    End Sub

    Function ValidateThis() As Boolean

        'true = everything OK

        ValidateThis = False

        Dim var1
        Dim strM As String

        var1 = Me.txtPN.Text
        If Len(var1) = 0 Then
            strM = "Permissions Name cannot be blank"
            GoTo end1
        End If

        'ensure name is unique
        Dim cbx As ComboBox = Me.cbxPermBase
        Dim Count1 As Int16
        Dim str1 As String
        For Count1 = 0 To cbx.Items.Count - 1

            str1 = cbx.Items(Count1).ToString
            If StrComp(str1, var1, CompareMethod.Text) = 0 Then
                strM = "The Permissions Group '" & var1 & "' already exists"
                GoTo end1
            End If

        Next

        ValidateThis = True

end1:
        If ValidateThis = False Then
            MsgBox(strM, vbInformation, "Invalid entry...")
        End If

    End Function

    Private Sub cbxPermBase_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cbxPermBase.SelectedIndexChanged


    End Sub
End Class