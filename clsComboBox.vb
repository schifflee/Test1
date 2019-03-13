
Public Class SelectedIndexChangingEventArgs
    Inherits EventArgs

    Public Property Cancel As Boolean = False
    Public Property NewIndex As Int32 = -1

    Friend Sub New(index As Int32)
        NewIndex = index
    End Sub
End Class


Public Class ComboBoxEx

    Inherits ComboBox

    Private selectedObject As Object = Nothing

    Public Event SelectedIndexChanging(sender As Object, e As SelectedIndexChangingEventArgs)

    Public Sub New()
        MyBase.New()
    End Sub

    Protected Overrides Sub OnSelectedIndexChanged(e As EventArgs)
        Dim evArgs As New SelectedIndexChangingEventArgs(MyBase.SelectedIndex)
        RaiseEvent SelectedIndexChanging(Me, evArgs)

        If evArgs.Cancel Then
            If selectedObject IsNot Nothing Then
                MyBase.SelectedIndex = MyBase.Items.IndexOf(selectedObject)
            Else
                MyBase.SelectedIndex = -1
            End If
            Return       ' do not fire Changed event
        End If

        MyBase.OnSelectedIndexChanged(e)
        selectedObject = MyBase.Items(MyBase.SelectedIndex)
    End Sub
End Class
