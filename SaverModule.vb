Module SQLModule

    Public Sub Saver(ctl As Object)

        'Get a generic command list first - Ignore errors (Multi table)
        Dim cb As New OleDb.OleDbCommandBuilder(Central.CurrentDataAdapter)

        Try
            Central.CurrentDataAdapter.UpdateCommand = cb.GetUpdateCommand()
        Catch
        End Try
        Try
            Central.CurrentDataAdapter.InsertCommand = cb.GetInsertCommand()
        Catch
        End Try
        Try
            Central.CurrentDataAdapter.DeleteCommand = cb.GetDeleteCommand()
        Catch
        End Try


        'Create and overwrite a custom one if needed (More than 1 table) ...OLEDB Parameters must be added in the order they are used
        Select Case ctl.name

            Case "DataGridView2", "DataGridView3"

                'SET THE Commands, with Parameters (OLDB Parameters must be added in the order they are used in the statement)
                Central.CurrentDataAdapter.UpdateCommand = New OleDb.OleDbCommand("UPDATE QueryCodes SET SiteCode=@P1, RespondCode=@P2, " & _
                                                                          "Person=@P3, TypeCode=@P4 " & _
                                                                          "WHERE QueryID=@P5")

                'Add parameters with the source columns in the dataset
                With Central.CurrentDataAdapter.UpdateCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.VarChar, 255, "SiteCode")
                    .Add("@P2", OleDb.OleDbType.VarChar, 255, "RespondCode")
                    .Add("@P3", OleDb.OleDbType.VarChar, 255, "Person")
                    .Add("@P4", OleDb.OleDbType.VarChar, 255, "TypeCode")
                    .Add("@P5", OleDb.OleDbType.Double, 255, "QueryID")
                End With


        End Select

        

        Call Central.SetCommandConnection()
        Call Central.UpdateBackend(ctl)

    End Sub



End Module
