Imports System.Data
Imports System.Linq
Imports System.Reflection

Public Class LinqUtilities

    'http://vbcity.com/blogs/mike-mcintyre/archive/2009/11/06/visual-basic-2008-create-a-datatable-from-linq-query-results.aspx

    'Friend Shared Function LINQToDataTable(Of T)(ByVal iEnumerableList As IEnumerable(Of T)) As DataTable
    Friend Shared Function LINQToDataTable(Of T)(ByVal iEnumerableList As IEnumerable) As DataTable
        Dim newDataTable As New DataTable()
        Dim thePropertyInfo As PropertyInfo() = Nothing
        If iEnumerableList Is Nothing Then
            Return newDataTable
        End If

        For Each item As T In iEnumerableList
            If thePropertyInfo Is Nothing Then
                thePropertyInfo = (DirectCast(item.[GetType](), Type)).GetProperties()
                For Each propInfo As PropertyInfo In thePropertyInfo
                    Dim columnDataType As Type = propInfo.PropertyType
                    If (columnDataType.IsGenericType) AndAlso (columnDataType.GetGenericTypeDefinition() Is GetType(Nullable(Of ))) Then
                        columnDataType = columnDataType.GetGenericArguments()(0)
                    End If
                    newDataTable.Columns.Add(New DataColumn(propInfo.Name, columnDataType))
                Next
            End If

            Dim dr As DataRow = newDataTable.NewRow()
            For Each pi As PropertyInfo In thePropertyInfo
                dr(pi.Name) = If(pi.GetValue(item, Nothing) Is Nothing, DBNull.Value, pi.GetValue(item, Nothing))
            Next

            newDataTable.Rows.Add(dr)
        Next
        Return newDataTable
    End Function

End Class

