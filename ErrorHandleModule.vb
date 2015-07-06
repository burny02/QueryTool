Module ErrorHandleModule
    Public Sub ErrorHandler(sender As Object, e As Object)

        Dim Obj As Object

        Try
            If TypeOf (sender) Is DataGridView Then
                Obj = CType(sender, DataGridView)
                Obj.Rows(e.RowIndex).Cells(e.ColumnIndex).ErrorText = e.exception.message
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Sub
End Module
