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

            Case "DataGridView2", "DataGridView3", "NewQueryGrid2", "NewQueryGrid3"

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

            Case "NewQueryGrid"

                Dim Status As String = "'Open'"
                Dim Study As String = "'" & AdQry.ComboBox2.SelectedValue & "'"
                Dim FieldName As String = "'Manual'"
                Dim CreateDate As String = "'" & Format(DateTime.Now, "dd-MMM-yyyy") & "'"
                Dim CreateTime As String = "'" & Format(DateTime.Now, "HH:mm") & "'"
                Dim CreatedBy As String = "'" & Overclass.GetUserName & "'"
                Dim CreatedByRole As String = "'" & Role & "'"
                Dim QueryID As String = "'MANUAL-" & _
                Overclass.TempDataTable("SELECT Count(QueryID) FROM Queries WHERE QueryID LIKE 'MANUAL-%'").Rows(0).Item(0) + 1 & "'"


                'SET THE Commands, with Parameters (OLDB Parameters must be added in the order they are used in the statement)
                Overclass.CurrentDataAdapter.InsertCommand = New OleDb.OleDbCommand("INSERT INTO Queries " & _
                "(QueryID, Study, RVLID, Initials, FormName, Status, PageNo, FieldName, Description, CreateDate, CreateTime, CreatedBy, CreatedByRole, VisitName) " & _
                "VALUES (" & QueryID & ", " & Study & ", @P1, @P2, @P3, " & Status & ", @P4, " & FieldName & ", @P5, " & CreateDate & ", " & CreateTime & ", " _
                & CreatedBy & ", " & CreatedByRole & ", @P6)")


                'Add parameters with the source columns in the dataset
                With Overclass.CurrentDataAdapter.InsertCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.VarChar, 255, "RVLID")
                    .Add("@P2", OleDb.OleDbType.VarChar, 255, "Initials")
                    .Add("@P3", OleDb.OleDbType.VarChar, 255, "FormName")
                    .Add("@P4", OleDb.OleDbType.VarChar, 255, "PageNo")
                    .Add("@P5", OleDb.OleDbType.VarChar, 255, "Description")
                    .Add("@P6", OleDb.OleDbType.VarChar, 255, "VisitName")
                End With

        End Select

        

        Call Overclass.SetCommandConnection()
        Call Overclass.UpdateBackend(ctl)

    End Sub



End Module
