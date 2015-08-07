Module SQLModule

    Public Sub Saver(ctl As Object)

        'Get a generic command list first - Ignore errors (Multi table)
        Dim cb As New OleDb.OleDbCommandBuilder(Overclass.CurrentDataAdapter)

        Try
            Overclass.CurrentDataAdapter.UpdateCommand = cb.GetUpdateCommand()
        Catch
        End Try
        Try
            Overclass.CurrentDataAdapter.InsertCommand = cb.GetInsertCommand()
        Catch
        End Try
        Try
            Overclass.CurrentDataAdapter.DeleteCommand = cb.GetDeleteCommand()
        Catch
        End Try


        'Create and overwrite a custom one if needed (More than 1 table) ...OLEDB Parameters must be added in the order they are used
        Select Case ctl.name

            Case "DataGridView2", "DataGridView3"

                'SET THE Commands, with Parameters (OLDB Parameters must be added in the order they are used in the statement)
                Overclass.CurrentDataAdapter.UpdateCommand = New OleDb.OleDbCommand("UPDATE QueryCodes SET SiteCode=@P1, RespondCode=@P2, " & _
                                                                          "Person=@P3, TypeCode=@P4 " & _
                                                                          "WHERE QueryID=@P5")

                'Add parameters with the source columns in the dataset
                With Overclass.CurrentDataAdapter.UpdateCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.VarChar, 5, "SiteCode")
                    .Add("@P2", OleDb.OleDbType.VarChar, 5, "RespondCode")
                    .Add("@P3", OleDb.OleDbType.VarChar, 5, "Person")
                    .Add("@P4", OleDb.OleDbType.VarChar, 5, "TypeCode")
                    .Add("@P5", OleDb.OleDbType.VarChar, 50, "QueryID")
                End With


        End Select

        

        Call Overclass.SetCommandConnection()
        Call Overclass.UpdateBackend(ctl)

    End Sub



End Module
