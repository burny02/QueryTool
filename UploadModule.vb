Module UploadModule

    Public Sub UploadCSV()

        MsgBox("Please ensure .xls contains all study queries")

        Dim fd As OpenFileDialog = New OpenFileDialog()

        Dim filnam As String
        Dim ColumnHeaderNumber As Integer = 2

        fd.Title = "Choose query .xls to upload"
        fd.InitialDirectory = "C:\"
        fd.Filter = ".xls|*.xls"
        fd.FilterIndex = 1
        fd.RestoreDirectory = True
        fd.Multiselect = False
        fd.ShowDialog()

        filnam = fd.FileName

        fd = Nothing

        If Not filnam = vbNullString Then

            Dim conStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filnam & ";Extended Properties='Excel 8.0;HDR=YES';"
            Dim connExcel As New OleDb.OleDbConnection(conStr)
            Dim cmdExcel As New OleDb.OleDbCommand()
            Dim oda As New OleDb.OleDbDataAdapter()
            Dim dt As New DataTable()

            cmdExcel.Connection = connExcel

            'Get the name of First Sheet 

            connExcel.Open()

            Dim dtExcelSchema As DataTable

            dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Tables, Nothing)

            Dim SheetName As String = dtExcelSchema.Rows(1)("TABLE_NAME").ToString()

            'Read Data from First Sheet 

            cmdExcel.CommandText = "SELECT [Protocol Number] AS F0, left([Screening Number],50) As F1, [Visit Name] AS F2, " & _
                                    "[Form Name] AS F3, left([Page Number],50) AS F4, " & _
                                    "[Field Name] AS F5, [Query Status] AS F6, replace([Query Text],""'"","""") AS F7, " & _
                                    "[Query Creation Date (UTC)] AS F8, [Query Creation Time (UTC)] AS F9, " & _
                                    "[Query Created By] AS F10, [Query Created By Role] AS F11, [Query Closed Date] AS F12, " & _
                                    "[Query Closed Time] AS F13, [Query Closed By] AS F14, [Query Closed By Role] AS F15, " & _
                                    "[Query ID] AS F16 " & _
                                    "FROM [" & SheetName & "]" & _
                                    "WHERE [Query ID]<>''"


            oda.SelectCommand = cmdExcel

            For Each column In dt.Columns
                column.DataType = GetType(String)
            Next
            oda.Fill(dt)

            Dim FinalCount As Long = (dt.Rows.Count)
            Dim Study As String = dt.Rows(1).Item(0)

            'Clean Up
            connExcel.Close()
            oda = Nothing
            connExcel = Nothing
            SheetName = Nothing
            dtExcelSchema = Nothing
            cmdExcel = Nothing
            conStr = Nothing


            'Get Backend
            Dim tempDa As OleDb.OleDbDataAdapter = Central.NewDataAdapter("SELECT * FROM Queries")
            Dim BackendDT As New DataTable()
            Dim UpdateTable = dt.Clone
            Dim InsertTable = dt.Clone
            tempDa.AcceptChangesDuringFill = True
            tempDa.Fill(BackendDT)



            'JOIN BACKEND AND EXCEL - PUT INTO UPDATE DATATABLE
            Dim UpdateData =
                    From a In dt.AsEnumerable()
                    Join b In BackendDT.AsEnumerable()
                    On
                       a.Field(Of String)("F16") Equals b.Field(Of String)("QueryID")
                    Where
                        a.Field(Of String)("F6") <> b.Field(Of String)("Status") Or
                        a.Field(Of String)("F12") <> b.Field(Of String)("ClosedDate") Or
                        a.Field(Of String)("F13") <> b.Field(Of String)("ClosedTime") Or
                        a.Field(Of String)("F14") <> b.Field(Of String)("ClosedBy") Or
                        a.Field(Of String)("F15") <> b.Field(Of String)("ClosedByRole")
                    Select a
            

            For Each row In UpdateData
                UpdateTable.ImportRow(row)
            Next


            UpdateData = Nothing

            'JOIN BACKEND AND EXCEL - DELETE QUERIES ALREADY ON
            Dim InsertData =
                    From a In dt.AsEnumerable()
                    Join b In BackendDT.AsEnumerable()
                    On
                        a.Field(Of String)("F16") Equals b.Field(Of String)("QueryID")
                Select a

            Dim RowsToDelete As New ArrayList()

            For Each row In InsertData
                RowsToDelete.Add(row)
            Next

            For Each row In RowsToDelete
                dt.Rows.Remove(row)
            Next

            For Each row In dt.Rows
                InsertTable.ImportRow(row)
            Next

            InsertData = Nothing



            'Clean Up
            RowsToDelete = Nothing
            tempDa = Nothing
            BackendDT = Nothing
            dt = Nothing



            'Update BackEnd
            Dim da As New OleDb.OleDbDataAdapter()


            'Set all rows as modified
            For Each row In UpdateTable.Rows
                row.SetModified()
            Next


            'Loop each record for update command

            For Each row In UpdateTable.Rows

                Dim P1 As String = "'" & Replace(row.Item("F6").ToString, "'", "") & "'"
                Dim P2 As String = "'" & Replace(row.Item("F12").ToString, "'", "") & "'"
                Dim P3 As String = "'" & Replace(row.Item("F13").ToString, "'", "") & "'"
                Dim P4 As String = "'" & Replace(row.Item("F14").ToString, "'", "") & "'"
                Dim P5 As String = "'" & Replace(row.Item("F15").ToString, "'", "") & "'"
                Dim P6 As String = "'" & Replace(row.Item("F16").ToString, "'", "") & "'"



                Call Central.ExecuteSQL("Update Queries SET  Status=" & P1 & ", ClosedDate=" & P2 & ", ClosedTime=" & P3 & _
                                        ", ClosedBy=" & P4 & ", ClosedByRole=" & P5 & _
                                        " WHERE QueryID=" & P6)

            Next



            'Set all rows as new
            For Each row In InsertTable.Rows
                row.SetAdded()
            Next

            'Set Insert Command for backend
            da.InsertCommand = New OleDb.OleDbCommand("INSERT INTO Queries " & _
                                                          "(Study, RVLID, VisitName, FormName, " & _
                                                          "PageNo, FieldName, Status, Description, CreateDate, " & _
                                                          "CreateTime, CreatedBy, CreatedByRole, ClosedDate, ClosedTime, " & _
                                                          "ClosedBy, ClosedByRole, QueryID) " & _
                                                      "VALUES (@P1, @P2, @P3, @P4, @P5, @P6, " & _
                                                          "@P7, @P8, @P9, @P10, @P11, @P12, @P13, @P14, @P15, " & _
                                                          "@P16, @P17)")


            'Set insert parameters
            With da.InsertCommand.Parameters
                .Add("@P1", OleDb.OleDbType.Char, 50, "F0")
                .Add("@P2", OleDb.OleDbType.Char, 50, "F1")
                .Add("@P3", OleDb.OleDbType.Char, 255, "F2")
                .Add("@P5", OleDb.OleDbType.Char, 255, "F3")
                .Add("@P4", OleDb.OleDbType.Char, 50, "F4")
                .Add("@P6", OleDb.OleDbType.Char, 50, "F5")
                .Add("@P7", OleDb.OleDbType.Char, 50, "F6")
                .Add("@P10", OleDb.OleDbType.Char, 255, "F7")
                .Add("@P11", OleDb.OleDbType.Char, 50, "F8")
                .Add("@P12", OleDb.OleDbType.Char, 50, "F9")
                .Add("@P13", OleDb.OleDbType.Char, 50, "F10")
                .Add("@P8", OleDb.OleDbType.Char, 50, "F11")
                .Add("@P9", OleDb.OleDbType.Char, 50, "F12")
                .Add("@P14", OleDb.OleDbType.Char, 50, "F13")
                .Add("@P15", OleDb.OleDbType.Char, 50, "F14")
                .Add("@P16", OleDb.OleDbType.Char, 50, "F15")
                .Add("@P17", OleDb.OleDbType.Char, 50, "F16")
            End With


            'Open Connection & Update
            Central.SetCommandConnection(da.InsertCommand)
            Central.OpenCon()
            da.Update(InsertTable)


            'Codes are entered as Data Macro


            'Close Off & Clean up
            Central.CloseCon()
            da = Nothing
            UpdateTable = Nothing
            InsertTable = Nothing
            dt = Nothing


            'Update Upload Date/Time
            Central.ExecuteSQL("UPDATE Study SET UploadDate=now(), UploadPerson='" & Central.GetUserName() & "'" & _
                       "WHERE StudyCode='" & Study & "'")


            'For QC check
            MsgBox("Upload Complete, " & FinalCount & " total queries uploaded")

            Call Central.Refresher(Form1.DataGridView1)

        End If

    End Sub


End Module
