'Imports needed for OLEDB connections (Access backend)
Imports System.Data
Imports System.Data.OleDb.OleDbConnection
Module SQLModule
    'Store the current dataset, adapter and binding source so that pubically accessable (Save operations etc) - Also then 1 main dataset per form view
    Public CurrentDataSet As DataSet = Nothing
    Public CurrentDataAdapter As OleDb.OleDbDataAdapter = Nothing
    Public CurrentBindingSource As BindingSource = Nothing
    'Connection information privately accessible 
    Private Const TablePath As String = "M:\VOLUNTEER SCREENING SERVICES\DavidBurnside\Queries\Backend3.accdb"
    Private Const PWord As String = "Crypto*Dave02"
    Private Const Connect As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & TablePath & ";Jet OLEDB:Database Password=" & PWord
    Public Const SolutionName As String = "Query Tool"

    Public Function QueryTest(SQLCode As String) As Long
        'Execute a SQL Command and return the number of records

        Dim Counter As Long
        Dim rs As New ADODB.Recordset

        Try
            'Connect
            rs.Open(SQLCode, Connect, ADODB.CursorTypeEnum.adOpenStatic)
            'Assign
            Counter = rs.RecordCount

        Catch ex As Exception
            MsgBox(ex.Message)

        Finally
            'Close Off & Clean up

            Try
                rs.Close()
            Catch ex As Exception
                ex = Nothing
            Finally
                rs = Nothing
            End Try

        End Try
        QueryTest = Counter

    End Function

    Public Sub ExecuteSQL(SQLCode As String)
        'Execute a SQL Command - No return

        'Create connection & Command
        Dim con As New OleDb.OleDbConnection(Connect)
        Dim cmd As New OleDb.OleDbCommand(SQLCode, con)

        Try
            'Open connection 
            con.Open()
            'Execute SQL Command
            cmd.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox(ex.Message)

        Finally

            'Close Off & Clean up
            Try
                con.Close()
            Catch ex As Exception
                ex = Nothing
            Finally
                con = Nothing
                cmd = Nothing
            End Try

        End Try

    End Sub

    Public Sub CreateDataSet(SQLCode As String, BindSource As BindingSource, ctl As Object)
        'Create a new dataset, set a bindining source and object to that binding source

        'Create Connection object
        Dim con As New OleDb.OleDbConnection(Connect)

        Try
            'Open connection
            con.Open()
            'Create New Dataset & adapter
            CurrentDataAdapter = New OleDb.OleDbDataAdapter(SQLCode, con)
            CurrentDataSet = New DataSet()
            CurrentBindingSource = BindSource

            'Use adapter to fill dataset
            CurrentDataAdapter.Fill(CurrentDataSet)

            'Set bindsource datasource as dataset, set object datasource as bindsource
            BindSource.DataSource = CurrentDataSet.Tables(0)
            ctl.datasource = BindSource

        Catch ex As Exception
            MsgBox(ex.Message)

        Finally

            'Close off & Clean up
            Try
                con.Close()
            Catch ex As Exception
                ex = Nothing
            Finally
                con = Nothing
            End Try

        End Try

    End Sub

    Public Sub UpdateBackend(ctl As Object)
        'Saving function to update access backend

        'New conecction 
        Dim con As New OleDb.OleDbConnection(Connect)

        'Set INSERT, UPDATE COMMANDS
        Call CustomCommand(ctl, con)


        'Is the data dirty / has errors that have auto-undone
        If CurrentDataSet.HasChanges() = False Then
            MsgBox("Errors present/No changes to upload")
            Exit Sub
        End If


        Try
            'Cancel current edit
            CurrentBindingSource.EndEdit()
            'Open connection
            con.Open()
            'Use dataadapter to update the backend (Commands already set)
            CurrentDataAdapter.Update(CurrentDataSet)
            MsgBox("Table Updated")
            'Remove any error messages & accept changes
            CurrentDataSet.AcceptChanges()

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally

            'Close off & clean up
            Try
                con.Close()
            Catch ex As Exception
                ex = Nothing
            Finally
                con = Nothing
            End Try

        End Try

    End Sub

    Public Function UnloadData() As Boolean
        'Close down currnt dataset, dataadapter & bindinsource

        'Variable if user wants to save
        Dim Cancel As Boolean = False


        'Is there currently a dataset to close?
        If IsNothing(CurrentDataSet) Then
            UnloadData = False
            Exit Function
        End If

        Try

            'Is the dataset dirty?
            If CurrentDataSet.HasChanges() Then

                'Ask user if they want to proceed and lose data?
                If (MsgBox("Changes to data will be lost unless saved first. Do you wish to discard changes?", vbYesNo) = vbNo) Then Cancel = True

            End If


            'If want to continue, clear all current data items
            If Cancel = False Then
                CurrentDataSet = Nothing
                CurrentDataAdapter = Nothing
                CurrentBindingSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            'Pass back whether clean up happened
            UnloadData = Cancel
        End Try

    End Function

    Public Sub CustomCommand(ctl As Object, connection As OleDb.OleDbConnection)
        'Create Custom INSERT, UPDATE COMMANDS for saving dataset (More than 1 table in select)


        Select Case ctl.name

            Case "DataGridView3"

                'Custom Command Builder...OLEDB Parameters must be added in the order they are used


                'New Connection
                Dim con As New OleDb.OleDbConnection(Connect)

                'SET THE Commands, with Parameters (OLDB Parameters must be added in the order they are used in the statement)
                CurrentDataAdapter.UpdateCommand = New OleDb.OleDbCommand("UPDATE TrainingCourse SET TypeID=@P1, CourseDate=@P2 WHERE ID=@P3", con)
                CurrentDataAdapter.InsertCommand = New OleDb.OleDbCommand("INSERT INTO TrainingCourse (TypeID, CourseDate) VALUES (@P1, @P2)", con)

                'Add parameters with the source columns in the dataset
                With CurrentDataAdapter.UpdateCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.Double).SourceColumn = "TypeID"
                    .Add("@P2", OleDb.OleDbType.Date).SourceColumn = "CourseDate"
                    .Add("@P3", OleDb.OleDbType.Double).SourceColumn = "ID"
                End With
                With CurrentDataAdapter.InsertCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.Double).SourceColumn = "TypeID"
                    .Add("@P2", OleDb.OleDbType.Date).SourceColumn = "CourseDate"
                End With


            Case Else

                'If not specified - Select commands with one table can auto generate INSERT, UPDATE commands
                Dim cb As New OleDb.OleDbCommandBuilder(CurrentDataAdapter)

        End Select

    End Sub

    Public Function TempDataSet(SQLCode As String) As DataSet
        'Create a temporary dataset for things such as combo box which arent based on the initial query

        'New connection
        Dim con As New OleDb.OleDbConnection(Connect)

        Try
            'Open connection
            con.Open()
            'New temporary data adapter and dataset
            Dim TempDataAdapter = New OleDb.OleDbDataAdapter(SQLCode, con)
            TempDataSet = New DataSet()
            'Use temp adapter to fill temp dataset
            TempDataAdapter.Fill(TempDataSet)

        Catch ex As Exception
            MsgBox(ex.Message)
            TempDataSet = Nothing
        Finally

            'Close off & Clean up
            Try
                con.Close()
            Catch ex As Exception
                ex = Nothing
            Finally
                con = Nothing
            End Try

        End Try

    End Function

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

            cmdExcel.CommandText = "SELECT first([Protocol Number]) AS F0, [Screening Number] As F1, [Visit Name] AS F2, " & _
                                    "[Form Name] AS F3, first([Page Number]) AS F4, " & _
                                    "first([Field Name]) AS F5, first([Query Status]) AS F6, [Query Text] AS F7, " & _
                                    "[Query Creation Date (UTC)] AS F8, [Query Creation Time (UTC)] AS F9, " & _
                                    "[Query Created By] AS F10, first([Query Created By Role]) AS F11, first([Query Closed Date]) AS F12, " & _
                                    "first([Query Closed Time]) AS F13, first([Query Closed By]) AS F14, first([Query Closed By Role]) AS F15 " & _
                                    "FROM [" & SheetName & "]" & _
                                    "GROUP BY [Screening Number], [Visit Name], [Form Name], [Query Text], [Query Creation Date (UTC)], " & _
                                    "[Query Creation Time (UTC)], [Query Created By]"


            oda.SelectCommand = cmdExcel

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
            Dim con As New OleDb.OleDbConnection(Connect)
            Dim tempDa As New OleDb.OleDbDataAdapter("SELECT * FROM Queries", con)
            Dim BackendDT As New DataTable()
            Dim UpdateTable = dt.Clone
            Dim InsertTable = dt.Clone
            tempDa.Fill(BackendDT)


            'JOIN BACKEND AND EXCEL - PUT INTO UPDATE DATATABLE
            Dim UpdateData =
                    From a In dt.AsEnumerable()
                    Join b In BackendDT.AsEnumerable()
                    On
                    a.Field(Of String)("F1") Equals b.Field(Of String)("RVLID") And
                    a.Field(Of String)("F2") Equals b.Field(Of String)("VisitName") And
                    a.Field(Of String)("F3") Equals b.Field(Of String)("FormName") And
                    a.Field(Of String)("F7") Equals b.Field(Of String)("Description") And
                    a.Field(Of String)("F8") Equals b.Field(Of String)("CreateDate") And
                    a.Field(Of String)("F9") Equals b.Field(Of String)("CreateTime") And
                    a.Field(Of String)("F10") Equals b.Field(Of String)("CreatedBy")
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
                    a.Field(Of String)("F1") Equals b.Field(Of String)("RVLID") And
                    a.Field(Of String)("F2") Equals b.Field(Of String)("VisitName") And
                    a.Field(Of String)("F3") Equals b.Field(Of String)("FormName") And
                    a.Field(Of String)("F7") Equals b.Field(Of String)("Description") And
                    a.Field(Of String)("F8") Equals b.Field(Of String)("CreateDate") And
                    a.Field(Of String)("F9") Equals b.Field(Of String)("CreateTime") And
                    a.Field(Of String)("F10") Equals b.Field(Of String)("CreatedBy")
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


            'Set Update Command for backend

            da.UpdateCommand = New OleDb.OleDbCommand("Update Queries SET Status=@P1, ClosedDate=@P2, ClosedTime=@P3, ClosedBy=@P4, ClosedByRole=@P5 " & _
                                                        "WHERE RVLID=@P6 AND VisitName=@P7 AND FormName=@P8 AND Description=@P9 " & _
                                                        "AND CreateDate=@P10 AND CreateTime=@P11 AND CreatedBy=@P12", con)

            'Set update parameters
            With da.UpdateCommand.Parameters
                .Add("@P1", OleDb.OleDbType.Char, 50, "F6")
                .Add("@P2", OleDb.OleDbType.Char, 50, "F12")
                .Add("@P3", OleDb.OleDbType.Char, 50, "F13")
                .Add("@P4", OleDb.OleDbType.Char, 50, "F14")
                .Add("@P5", OleDb.OleDbType.Char, 50, "F15")
                .Add("@P6", OleDb.OleDbType.Char, 50, "F1")
                .Add("@P7", OleDb.OleDbType.Char, 50, "F2")
                .Add("@P8", OleDb.OleDbType.Char, 50, "F3")
                .Add("@P9", OleDb.OleDbType.Char, 50, "F7")
                .Add("@P10", OleDb.OleDbType.Char, 50, "F8")
                .Add("@P11", OleDb.OleDbType.Char, 50, "F9")
                .Add("@P12", OleDb.OleDbType.Char, 50, "F10")
            End With

            con.Open()
            da.Update(UpdateTable)
            con.Close()


            'Set all rows as new
            For Each row In InsertTable.Rows
                row.SetAdded()
            Next

            'Set Insert Command for backend
            da.InsertCommand = New OleDb.OleDbCommand("INSERT INTO Queries " & _
                                                          "(Study, RVLID, VisitName, FormName, " & _
                                                          "PageNo, FieldName, Status, Description, CreateDate, " & _
                                                          "CreateTime, CreatedBy, CreatedByRole, ClosedDate, ClosedTime, " & _
                                                          "ClosedBy, ClosedByRole) " & _
                                                      "VALUES (@P1, @P2, @P3, @P4, @P5, @P6, " & _
                                                          "@P7, @P8, @P9, @P10, @P11, @P12, @P13, @P14, @P15, " & _
                                                          "@P16)", con)


            'Set insert parameters
            With da.InsertCommand.Parameters
                .Add("@P1", OleDb.OleDbType.Char, 50, "F0")
                .Add("@P2", OleDb.OleDbType.Char, 50, "F1")
                .Add("@P3", OleDb.OleDbType.Char, 100, "F2")
                .Add("@P5", OleDb.OleDbType.Char, 100, "F3")
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
                .Add("@P15", OleDb.OleDbType.Char, 30, "F14")
                .Add("@P16", OleDb.OleDbType.Char, 30, "F15")
            End With


            'Open Connection & Update
            con.Open()
            da.Update(InsertTable)


            'Codes are entered as Data Macro


            'Close Off & Clean up
            con.Close()
            con = Nothing
            da = Nothing
            UpdateTable = Nothing
            InsertTable = Nothing
            dt = Nothing


            'Update Upload Date/Time
            ExecuteSQL("UPDATE Study SET UploadDate=Date(), UploadPerson='" & GetUserName() & "'" & _
                       "WHERE StudyCode='" & Study & "'")


            'For QC check
            MsgBox("Upload Complete, " & FinalCount & " total queries uploaded")

            Form1.TabControl1.SelectTab(0)

        End If

    End Sub
End Module
