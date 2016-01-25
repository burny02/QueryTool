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

            Case "DataGridView2", "NewQueryGrid2", "NewQueryGrid3"

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
                Dim Study As String = "'" & AdQry.ComboBox101.SelectedValue & "'"
                Dim FieldName As String = "'Manual'"
                Dim CreateDate As String = "'" & Format(DateTime.Now, "dd-MMM-yyyy") & "'"
                Dim CreateTime As String = "'" & Format(DateTime.Now, "HH:mm") & "'"
                Dim CreatedBy As String = "'" & Overclass.GetUserName & "'"
                Dim CreatedByRole As String = "'" & Role & "'"
                Dim PassNo As Double = 0

                Overclass.CmdList.Clear()

                For Each row In Overclass.CurrentDataSet.Tables(0).Rows

                    If row.RowState = DataRowState.Added Then

                        PassNo = PassNo + 1

                        If IsDBNull(row.item("RVLID")) Then
                            MsgBox("RVL ID missing")
                            Overclass.CmdList.Clear()
                            Exit Sub
                        End If

                        If IsDBNull(row.item("Initials")) Then
                            MsgBox("Initials missing")
                            Overclass.CmdList.Clear()
                            Exit Sub
                        End If

                        If IsDBNull(row.item("FormName")) Then
                            MsgBox("Form Name missing")
                            Overclass.CmdList.Clear()
                            Exit Sub
                        End If

                        If IsDBNull(row.item("PageNo")) Then
                            MsgBox("Page No missing")
                            Overclass.CmdList.Clear()
                            Exit Sub
                        End If

                        If IsDBNull(row.item("Description")) Then
                            MsgBox("Description missing")
                            Overclass.CmdList.Clear()
                            Exit Sub
                        End If

                        If IsDBNull(row.item("VisitName")) Then
                            MsgBox("Visit Name missing")
                            Overclass.CmdList.Clear()
                            Exit Sub
                        End If

                        Dim RVLID As String = "'" & row.item("RVLID") & "'"
                        Dim Initials As String = "'" & row.item("Initials") & "'"
                        Dim FormName As String = "'" & row.item("FormName") & "'"
                        Dim PageNo As String = "'" & row.item("PageNo") & "'"
                        Dim Description As String = "'" & row.item("Description") & "'"
                        Dim VisitName As String = "'" & row.item("VisitName") & "'"

                        Dim QueryID As String = "'MANUAL-" & _
                        Overclass.TempDataTable("SELECT Count(QueryID) FROM Queries WHERE QueryID LIKE 'MANUAL-%'").Rows(0).Item(0) + PassNo & "'"

                        Dim InsertCmd As OleDb.OleDbCommand

                        'SET THE Commands, with Parameters (OLDB Parameters must be added in the order they are used in the statement)
                        InsertCmd = New OleDb.OleDbCommand("INSERT INTO Queries " & _
                        "(QueryID, Study, RVLID, Initials, FormName, Status, PageNo, FieldName, Description, CreateDate, CreateTime, CreatedBy, CreatedByRole, VisitName) " & _
                        "VALUES (" & QueryID & "," & Study & ", " & RVLID & "," & Initials & "," & FormName & "," & Status & "," & PageNo & "," & FieldName & "," & Description & _
                        "," & CreateDate & "," & CreateTime & "," & CreatedBy & "," & CreatedByRole & "," & VisitName & ")")


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


        End Select

        

        Call Overclass.SetCommandConnection()
        Call Overclass.UpdateBackend(ctl, SaveMessage)

    End Sub



End Module
