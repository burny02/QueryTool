Module SQLModule

    Public Sub Saver(ctl As Object)

        Dim SaveMessage As Boolean = True

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

            Case "NewQueryGrid"

                Dim Status As String = "'Open'"
                Dim CreateDate As String = "'" & Format(DateTime.Now, "dd-MMM-yyyy") & "'"
                Dim CreateTime As String = "'" & Format(DateTime.Now, "HH:mm") & "'"
                Dim CreatedBy As String = "'" & Overclass.GetUserName & "'"
                Dim CreatedByRole As String = "'" & Role & "'"
                Dim Combine As String = Status & "," & CreateDate & "," & CreateTime & "," & CreatedBy & "," & CreatedByRole

                Overclass.CurrentDataAdapter.InsertCommand = New OleDb.OleDbCommand("INSERT INTO Queries " &
                "(Status, CreateDate, CreateTime, CreatedBy, CreatedByRole, Study, RVLID, " &
                "Initials, FormName, PageNo, Description, VisitName, Priority, SiteCode, RespondCode, Person, TypeCode, AssCode, QName) " &
                "VALUES (" & Combine & ",@P0, @P1, @P2, @P3, @P4, @P5, @P6, @P7, @P8, @P9, @P10, @P11, @P12, @P13)")

                With Overclass.CurrentDataAdapter.InsertCommand.Parameters
                    .Add("@P0", OleDb.OleDbType.VarChar, 50, "Study")
                    .Add("@P1", OleDb.OleDbType.VarChar, 50, "RVLID")
                    .Add("@P2", OleDb.OleDbType.VarChar, 50, "Initials")
                    .Add("@P3", OleDb.OleDbType.VarChar, 255, "FormName")
                    .Add("@P4", OleDb.OleDbType.VarChar, 50, "PageNo")
                    .Add("@P5", OleDb.OleDbType.LongVarChar, 500, "Description")
                    .Add("@P6", OleDb.OleDbType.VarChar, 255, "VisitName")
                    .Add("@P7", OleDb.OleDbType.Boolean, 50, "Priority")
                    .Add("@P8", OleDb.OleDbType.VarChar, 3, "SiteCode")
                    .Add("@P9", OleDb.OleDbType.VarChar, 3, "RespondCode")
                    .Add("@P10", OleDb.OleDbType.VarChar, 3, "Person")
                    .Add("@P11", OleDb.OleDbType.VarChar, 3, "TypeCode")
                    .Add("@P12", OleDb.OleDbType.VarChar, 3, "AssCode")
                    .Add("@P13", OleDb.OleDbType.VarChar, 5, "QName")
                End With

                Overclass.CurrentDataAdapter.UpdateCommand = New OleDb.OleDbCommand("UPDATE Queries SET  
                Status=@P1, ClosedDate=@P3, ClosedTime=@P4, ClosedBy=@P5, ClosedByRole=@P6, 
                RVLID=@P7, Initials=@P8, VisitName=@P9, FormName=@10, PageNo=@P11, Description=@P12, 
                Priority=@P14, Bounced=@P15, SiteCode=@P16, RespondCode=@P17, Person=@P18, TypeCode=@P19, AssCode=@P20, PDFLink=@P21, QName=@P22
                WHERE QueryID=@P23")

                'Add parameters with the source columns in the dataset
                With Overclass.CurrentDataAdapter.UpdateCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.VarChar, 50, "Status")
                    .Add("@P3", OleDb.OleDbType.VarChar, 50, "ClosedDate")
                    .Add("@P4", OleDb.OleDbType.VarChar, 50, "ClosedTime")
                    .Add("@P5", OleDb.OleDbType.VarChar, 50, "ClosedBy")
                    .Add("@P6", OleDb.OleDbType.VarChar, 50, "ClosedByRole")
                    .Add("@P7", OleDb.OleDbType.VarChar, 50, "RVLID")
                    .Add("@P8", OleDb.OleDbType.VarChar, 50, "Initials")
                    .Add("@P9", OleDb.OleDbType.VarChar, 255, "VisitName")
                    .Add("@P10", OleDb.OleDbType.VarChar, 255, "FormName")
                    .Add("@P11", OleDb.OleDbType.VarChar, 50, "PageNo")
                    .Add("@P12", OleDb.OleDbType.LongVarChar, 500, "Description")
                    .Add("@P14", OleDb.OleDbType.Boolean, 50, "Priority")
                    .Add("@P15", OleDb.OleDbType.Boolean, 50, "Bounced")
                    .Add("@P16", OleDb.OleDbType.VarChar, 3, "SiteCode")
                    .Add("@P17", OleDb.OleDbType.VarChar, 3, "RespondCode")
                    .Add("@P18", OleDb.OleDbType.VarChar, 3, "Person")
                    .Add("@P19", OleDb.OleDbType.VarChar, 3, "TypeCode")
                    .Add("@P20", OleDb.OleDbType.VarChar, 3, "AssCode")
                    .Add("@P21", OleDb.OleDbType.VarChar, 255, "PDFLink")
                    .Add("@P22", OleDb.OleDbType.VarChar, 5, "QName")
                    .Add("@P23", OleDb.OleDbType.VarChar, 50, "QueryID")
                End With



        End Select



        Call Overclass.SetCommandConnection()
        Call Overclass.UpdateBackend(ctl, SaveMessage)
        For Each cmd As OleDb.OleDbCommand In RespondCommands
            Overclass.AddToMassSQL(cmd)
        Next
        Overclass.ExecuteMassSQL()
        RespondCommands.Clear()

    End Sub



End Module
