Partial Class StudyDoc_01DataSet1
    Partial Class TBLDATADataTable

        Private Sub TBLDATADataTable_TBLDATARowChanging(sender As Object, e As TBLDATARowChangeEvent) Handles Me.TBLDATARowChanging

        End Sub

    End Class

    Partial Class TBLASSIGNEDSAMPLESDataTable

        Private Sub TBLASSIGNEDSAMPLESDataTable_ColumnChanging(sender As Object, e As DataColumnChangeEventArgs) Handles Me.ColumnChanging
            If (e.Column.ColumnName = Me.BOOLINTSTDColumn.ColumnName) Then
                'Add user code here
            End If

        End Sub

    End Class

End Class

Namespace StudyDoc_01DataSet1TableAdapters

    Partial Public Class TBLASSIGNEDSAMPLESTableAdapter
    End Class
End Namespace
