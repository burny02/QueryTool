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

            Case "DataGridView2", "NewQueryGrid2"

                'SET THE Commands, with Parameters (OLDB Parameters must be added in the order they are used in the statement)
                Overclass.CurrentDataAdapter.UpdateCommand = New OleDb.OleDbCommand("UPDATE QueryCodes SET SiteCode=@P1, RespondCode=@P2, " &
                                                                          "Person=@P3, TypeCode=@P4 " &
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
                Dim FieldName As String = "'Manual'"
                Dim CreateDate As String = "'" & Format(DateTime.Now, "dd-MMM-yyyy") & "'"
                Dim CreateTime As String = "'" & Format(DateTime.Now, "HH:mm") & "'"
                Dim CreatedBy As String = "'" & Overclass.GetUserName & "'"
                Dim CreatedByRole As String = "'" & Role & "'"
                Dim PassNo As Double = 0

                Overclass.CmdList.Clear()

                For Each row In Overclass.CurrentDataSet.Tables(0).Rows

                    If row.RowState = DataRowState.Added Then

                        If String.IsNullOrWhiteSpace(row.item("RVLID").ToString) And
                        String.IsNullOrWhiteSpace(row.item("Initials").ToString) And
                        String.IsNullOrWhiteSpace(row.item("FormName").ToString) And
                        String.IsNullOrWhiteSpace(row.item("PageNo").ToString) And
                        String.IsNullOrWhiteSpace(row.item("Description").ToString) And
                        String.IsNullOrWhiteSpace(row.item("VisitName").ToString) Then

                            Continue For

                        End If

                        PassNo = PassNo + 1

                        If String.IsNullOrWhiteSpace(row.item("RVLID").ToString) Then
                            MsgBox("RVL ID missing")
                            Overclass.CmdList.Clear()
                            Exit Sub
                        End If

                        If String.IsNullOrWhiteSpace(row.item("Initials").ToString) Then
                            MsgBox("Initials missing")
                            Overclass.CmdList.Clear()
                            Exit Sub
                        End If

                        If String.IsNullOrWhiteSpace(row.item("FormName").ToString) Then
                            MsgBox("Form Name missing")
                            Overclass.CmdList.Clear()
                            Exit Sub
                        End If

                        If String.IsNullOrWhiteSpace(row.item("PageNo").ToString) Then
                            MsgBox("Page No missing")
                            Overclass.CmdList.Clear()
                            Exit Sub
                        End If

                        If String.IsNullOrWhiteSpace(row.item("Description").ToString) Then
                            MsgBox("Description missing")
                            Overclass.CmdList.Clear()
                            Exit Sub
                        End If

                        If String.IsNullOrWhiteSpace(row.item("VisitName").ToString) Then
                            MsgBox("Visit Name missing")
                            Overclass.CmdList.Clear()
                            Exit Sub
                        End If

                        If String.IsNullOrWhiteSpace(row.item("Priority").ToString) Then
                            MsgBox("Priority missing")
                            Overclass.CmdList.Clear()
                            Exit Sub
                        End If

                        Dim RVLID As String = "'" & row.item("RVLID") & "'"
                            Dim Initials As String = "'" & row.item("Initials") & "'"
                            Dim FormName As String = "'" & row.item("FormName") & "'"
                            Dim PageNo As String = "'" & row.item("PageNo") & "'"
                            Dim Description As String = "'" & row.item("Description") & "'"
                            Dim VisitName As String = "'" & row.item("VisitName") & "'"
                            Dim Study As String = "'" & row.item("Study") & "'"
                            Dim ResolvedBy As String = "'" & row.item("ResolvedBy") & "'"
                            Dim ResolvedDate As String = "'" & row.item("ResolvedDate") & "'"
                            Dim Priority As String = "'" & row.item("Priority") & "'"

                            Dim QueryID As String = "'MANUAL-" &
                        Overclass.TempDataTable("SELECT Max(CLng(Replace([QueryID],'Manual-',''))) AS WhatNo FROM (SELECT Queries.QueryID " &
                        "FROM Queries Where (((Queries.QueryID) Like 'MANUAL-%')))  AS a").Rows(0).Item(0) + PassNo & "'"

                            Dim InsertCmd As OleDb.OleDbCommand

                            'SET THE Commands, with Parameters (OLDB Parameters must be added in the order they are used in the statement)
                            InsertCmd = New OleDb.OleDbCommand("INSERT INTO Queries " &
                        "(QueryID, Study, RVLID, Initials, FormName, Status, PageNo, FieldName, Description, CreateDate, CreateTime, CreatedBy, CreatedByRole, VisitName, " &
                        "Priority, ResolvedBy, ResolvedDate) " &
                        "VALUES (" & QueryID & "," & Study & ", " & RVLID & "," & Initials & "," & FormName & "," & Status & "," & PageNo & "," & FieldName & "," & Description &
                        "," & CreateDate & "," & CreateTime & "," & CreatedBy & "," & CreatedByRole & "," & VisitName & "," & Priority & "," & ResolvedBy & "," & ResolvedDate & ")")


                            Overclass.AddToMassSQL(InsertCmd)

                        End If

                Next

                Try
                    Overclass.ExecuteMassSQL()
                    SaveMessage = False
                Catch ex As Exception
                    MsgBox(ex.Message)
                    Exit Sub

                End Try

                For Each row As DataRow In Overclass.CurrentDataSet.Tables(0).Rows

                    If row.RowState = DataRowState.Added Then
                        row.AcceptChanges()
                        SaveMessage = False
                    End If

                Next

                MsgBox("Table Updated")

                Overclass.CurrentDataAdapter.UpdateCommand = New OleDb.OleDbCommand("UPDATE Queries SET  
                Status=@P1, FieldName=@P2, ClosedDate=@P3, 
                ClosedTime=@P4, ClosedBy=@P5, ClosedByRole=@P6, 
                RVLID=@P7, Initials=@P8, VisitName=@P9, 
                FormName=@10, PageNo=@P11, Description=@P12, 
                Priority=@P14, ResolvedBy=@P15, ResolvedDate=@P16
                WHERE QueryID=@P17")

                'Add parameters with the source columns in the dataset
                With Overclass.CurrentDataAdapter.UpdateCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.VarChar, 50, "Status")
                    .Add("@P2", OleDb.OleDbType.VarChar, 255, "FieldName")
                    .Add("@P3", OleDb.OleDbType.VarChar, 50, "ClosedDate")
                    .Add("@P4", OleDb.OleDbType.VarChar, 50, "ClosedTime")
                    .Add("@P5", OleDb.OleDbType.VarChar, 50, "ClosedBy")
                    .Add("@P6", OleDb.OleDbType.VarChar, 50, "ClosedByRole")
                    .Add("@P7", OleDb.OleDbType.VarChar, 50, "RVLID")
                    .Add("@P8", OleDb.OleDbType.VarChar, 50, "Initials")
                    .Add("@P9", OleDb.OleDbType.VarChar, 255, "VisitName")
                    .Add("@P10", OleDb.OleDbType.VarChar, 255, "FormName")
                    .Add("@P11", OleDb.OleDbType.VarChar, 50, "PageNo")
                    .Add("@P12", OleDb.OleDbType.VarChar, 255, "Description")
                    .Add("@P14", OleDb.OleDbType.VarChar, 50, "Priority")
                    .Add("@P15", OleDb.OleDbType.VarChar, 50, "ResolvedBy")
                    .Add("@P16", OleDb.OleDbType.VarChar, 50, "ResolvedDate")
                    .Add("@P17", OleDb.OleDbType.VarChar, 50, "QueryID")
                End With

        End Select



        Call Overclass.SetCommandConnection()
        Call Overclass.UpdateBackend(ctl, SaveMessage)

    End Sub



End Module
