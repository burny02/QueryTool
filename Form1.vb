Public Class Form1
    Private AssTable As DataTable

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Visible = False
        Me.Hide()

        Call StartUpCentral()

        If AccessLevel = 0 Then

            RespView = New ResponseView
            Specifics(RespView)

            RespView.FilterCombo1.LiveData = False
            RespView.FilterCombo1.SetAsExternalSource("SiteCode", "Site", "SELECT DISTINCT Code AS SiteCode, Site FROM SiteCode", Overclass)
            RespView.FilterCombo2.LiveData = False
            RespView.FilterCombo2.SetAsExternalSource("RespondCode", "Group", "SELECT DISTINCT Code AS RespondCode, Group FROM GroupCode", Overclass)
            RespView.StaffQueryGrid.Columns("QueryID").Visible = False

            RespView.Text = SolutionName
            RespView.ShowDialog()
            Application.Exit()

        Else

            Me.Visible = True
            Me.WindowState = FormWindowState.Maximized

        End If


        Try
            Me.Label2.Text = "Query Tool " & vbNewLine & "Developed by David Burnside" & vbNewLine & "Version:     " & System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
        Catch ex As Exception
            Me.Label2.Text = "Query Tool " & vbNewLine & "Developed by David Burnside"
        End Try

        Me.Text = SolutionName

    End Sub

    Private Sub TabControl1_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabControl1.Selecting

        Dim SQLCode As String = vbNullString
        Dim Bind As BindingSource = BindingSource1
        Dim ctl As Object = Nothing

        If Overclass.UnloadData() = True Then
            e.Cancel = True
            Exit Sub
        End If

        Overclass.ResetCollection()
        RespondCommands.Clear()

        Select Case e.TabPage.Text


            Case "Add Queries"
                NewQueryGrid.Columns.Clear()
                Specifics(NewQueryGrid)

            Case "Reports"
                Me.DateTimePicker2.Value = Date.Now

        End Select


        Call Specifics(ctl)

    End Sub


    Public Sub Specifics(ctl As Object)

        If IsNothing(ctl) Then Exit Sub

        Dim SqlCode As String = vbNullString

        Select Case ctl.name

            Case "ResponseView"
                Dim CurrentFilter As String = ""
                Dim Filter1 As String = ""
                Dim Filter2 As String = ""
                Dim Filter3 As String = ""
                Dim Filter4 As String = ""
                Dim Filter5 As String = ""

                Try
                    CurrentFilter = (Overclass.CurrentDataSet.Tables(0).DefaultView.RowFilter)
                Catch ex As Exception
                End Try

                Try
                    Filter1 = RespView.FilterCombo1.Text
                Catch ex As Exception
                End Try

                Try
                    Filter2 = RespView.FilterCombo2.Text
                Catch ex As Exception
                End Try

                Try
                    Filter3 = RespView.FilterCombo3.Text
                Catch ex As Exception
                End Try

                Try
                    Filter4 = RespView.FilterCombo30.Text
                Catch ex As Exception
                End Try

                Try
                    Filter5 = RespView.FilterCombo90.Text
                Catch ex As Exception
                End Try

                RespView.StaffQueryGrid.Columns.Clear()
                SqlCode = "SELECT CreatedByRole, QueryID, Study, Person, " &
                "replace(replace(cstr(Priority),0,'False'),-1,'True') AS PRI, Initials & ' ' & RVLID AS Volunteer, VisitName, FormName, PageNo, Description, SiteCode, RespondCode, Bounced " &
                "FROM Queries INNER JOIN Study ON Queries.Study=Study.StudyCode WHERE Hidden=False AND Status ='Open' ORDER BY CreateDate ASC"
                Overclass.CreateDataSet(SqlCode, RespView.BindingSource1, RespView.StaffQueryGrid)

                With RespView.StaffQueryGrid
                    .ReadOnly = True
                    .Columns("QueryID").Visible = False
                    .Columns("Bounced").Visible = False
                    .Columns("Study").Visible = False
                    .Columns("SiteCode").Visible = False
                    .Columns("RespondCode").Visible = False
                    .Columns("PRI").Visible = False
                    .Columns("VisitName").HeaderText = "Study Visit"
                    .Columns("FormName").HeaderText = "Assessment/Procedure"
                    .Columns("PageNo").HeaderText = "Page No"
                    .Columns("Person").HeaderText = "Assigned"
                    .Columns("CreatedByRole").HeaderText = "Raised By"
                    Dim clm2 As New DataGridViewImageColumn
                    clm2.HeaderText = "Respond"
                    clm2.Name = "RespondClm"
                    clm2.ImageLayout = DataGridViewImageCellLayout.Zoom
                    clm2.Image = My.Resources.speech
                    .Columns.Add(clm2)
                    .Columns("RespondClm").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                    .Columns("PageNo").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                    .Columns("Person").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                    .Columns("Volunteer").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                    .Columns("CreatedByRole").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                End With

                With RespView
                    .FilterCombo4.SetAsInternalSource("CreatedByRole", "CreatedByRole", Overclass)
                    .FilterCombo3.AllowBlanks = False
                    .FilterCombo3.SetAsInternalSource("Study", "Study", Overclass)
                    .FilterCombo30.SetAsInternalSource("PRI", "PRI", Overclass)
                    .FilterCombo90.SetAsInternalSource("Volunteer", "Volunteer", Overclass)

                    .StaffQueryGrid.Columns("QueryID").Visible = False
                End With

                If CurrentFilter <> "" Then Overclass.CurrentDataSet.Tables(0).DefaultView.RowFilter = CurrentFilter
                RespView.FilterCombo1.Text = Filter1
                RespView.FilterCombo2.Text = Filter2
                RespView.FilterCombo3.Text = Filter3
                RespView.FilterCombo30.Text = Filter4
                RespView.FilterCombo90.Text = Filter5

            Case "NewQueryGrid"

                NewQueryGrid.Columns.Clear()

                SqlCode = "SELECT CreatedByRole, SiteCode, RespondCode, Person, TypeCode, Status, Study, QueryID, CreatedBy, CreateDate, CreateTime, " &
                    "ClosedDate, ClosedTime, ClosedBy, ClosedByRole, RVLID, Initials, " &
                    "VisitName, FormName, PageNo, Description, Priority, Bounced, AssCode, PDFLink " &
                    "FROM Queries INNER JOIN Study On queries.Study=Study.StudyCode " &
                    "WHERE Hidden=false " &
                    "AND Status<>'Closed' " &
                    "ORDER BY Status DESC, RVLID ASC"

                NewQueryGrid.AutoGenerateColumns = True
                Overclass.CreateDataSet(SqlCode, BindingSource1, NewQueryGrid)
                NewQueryGrid.AutoGenerateColumns = False

                FilterCombo1.SetAsInternalSource("CreatedByRole", "CreatedByRole", Overclass)
                FilterCombo1.SetDGVDefault(ctl, "CreatedByRole")
                FilterCombo7.LiveData = False
                FilterCombo7.SetAsExternalSource("SiteCode", "Site", "SELECT DISTINCT Code AS SiteCode, Site FROM SiteCode", Overclass)
                FilterCombo30.AllowBlanks = False
                FilterCombo30.SetAsExternalSource("Study", "Study", "SELECT StudyCode As Study FROM Study " &
                "WHERE Hidden=False ORDER BY StudyCode ASC", Overclass)
                FilterCombo30.SetDGVDefault(ctl, "Study")

                FilterCombo90.SetAsInternalSource("Initials", "Initials", Overclass)
                FilterCombo100.SetAsInternalSource("Status", "Status", Overclass)

                FilterCombo20.SetAsInternalSource("RVLID", "RVLID", Overclass)
                FilterCombo10.SetAsInternalSource("VisitName", "VisitName", Overclass)

                NewQueryGrid.Columns("Bounced").Visible = False
                NewQueryGrid.Columns("QueryID").Visible = False
                NewQueryGrid.Columns("Study").Visible = False
                NewQueryGrid.Columns("CreateDate").Visible = False
                NewQueryGrid.Columns("CreateTime").Visible = False
                NewQueryGrid.Columns("ClosedDate").Visible = False
                NewQueryGrid.Columns("ClosedTime").Visible = False
                NewQueryGrid.Columns("ClosedBy").Visible = False
                NewQueryGrid.Columns("ClosedByRole").Visible = False
                NewQueryGrid.Columns("Study").Visible = False
                NewQueryGrid.Columns("Status").Visible = False
                NewQueryGrid.Columns("AssCode").Visible = False
                NewQueryGrid.Columns("PDFLink").Visible = False

                NewQueryGrid.Columns("CreatedBy").ReadOnly = True

                NewQueryGrid.Columns("RVLID").HeaderText = "Subject"
                NewQueryGrid.Columns("VisitName").HeaderText = "Study Visit"
                NewQueryGrid.Columns("FormName").HeaderText = "Assessment/Procedure"
                NewQueryGrid.Columns("PageNo").HeaderText = "Page"
                NewQueryGrid.Columns("Description").HeaderText = "Query Description"
                NewQueryGrid.Columns("CreatedBy").HeaderText = "Created By"
                NewQueryGrid.Columns("CreatedByRole").HeaderText = "Role"




                Dim clm3 As DataGridViewColumn = Overclass.SetUpNewComboColumn("SELECT Site AS Display, Code FROM SiteCode " &
                                                   "a inner join Study b ON a.ListID=b.CodeList " &
                                                   "WHERE CStr(StudyCode)=", FilterCombo30,
                                                  "Code", "Display", "SiteCode", "Site", NewQueryGrid, "SiteClm")

                Dim clm4 As DataGridViewColumn = Overclass.SetUpNewComboColumn("SELECT Group AS Display, Code FROM GroupCode " &
                                                   "a inner join Study b ON a.ListID=b.CodeList " &
                                                   "WHERE CStr(StudyCode)=", FilterCombo30,
                                                  "Code", "Display", "RespondCode", "Group", NewQueryGrid, "GroupClm")


                Dim clm5 As DataGridViewColumn = Overclass.SetUpNewComboColumn("SELECT ErrorType AS Display, Code FROM TypeCode " &
                                                   "a inner join Study b ON a.ListID=b.CodeList " &
                                                   "WHERE CStr(StudyCode)=", FilterCombo30,
                                                  "Code", "Display", "TypeCode", "Type", NewQueryGrid, "TypeClm")

                Dim clm6 As New DataGridViewComboBoxColumn
                If AssTable Is Nothing Then AssTable = Overclass.TempDataTable("SELECT AssName, AssCode From AssType ORDER BY AssName")
                clm6.DataSource = AssTable
                clm6.DisplayMember = "AssName"
                clm6.ValueMember = "AssCode"
                clm6.DataPropertyName = "AssCode"
                clm6.Name = "AssDrop"
                clm6.HeaderText = "Assessment Type"

                NewQueryGrid.Columns.Add(clm6)

                Dim cmb2 As New DataGridViewImageColumn
                cmb2.HeaderText = "Copy"
                cmb2.Image = My.Resources.copy
                cmb2.ImageLayout = DataGridViewImageCellLayout.Zoom
                cmb2.Name = "CopyQuery"

                NewQueryGrid.Columns.Add(cmb2)


                Dim clm2 As New DataGridViewImageColumn
                clm2.HeaderText = "Respond"
                clm2.Name = "RespondClm"
                clm2.ImageLayout = DataGridViewImageCellLayout.Zoom
                clm2.Image = My.Resources.speech
                NewQueryGrid.Columns.Add(clm2)


                Dim cmb As New DataGridViewImageColumn
                cmb.HeaderText = "Close"
                cmb.Image = My.Resources.TICK
                cmb.ImageLayout = DataGridViewImageCellLayout.Zoom
                cmb.Name = "StatusCmb"

                NewQueryGrid.Columns.Add(cmb)

                Dim pdfClm As New DataGridViewImageColumn
                pdfClm.HeaderText = "PDF"
                pdfClm.Name = "PDF"
                pdfClm.ImageLayout = DataGridViewImageCellLayout.Zoom
                pdfClm.Image = My.Resources.PDF
                NewQueryGrid.Columns.Add(pdfClm)


                NewQueryGrid.Columns("SiteCode").Visible = False
                NewQueryGrid.Columns("CreatedBy").Visible = False
                NewQueryGrid.Columns("RespondCode").Visible = False
                NewQueryGrid.Columns("TypeCode").Visible = False

                'Visible
                NewQueryGrid.Columns("CreatedByRole").DisplayIndex = 0
                NewQueryGrid.Columns("RVLID").DisplayIndex = 1
                NewQueryGrid.Columns("Initials").DisplayIndex = 2
                NewQueryGrid.Columns("VisitName").DisplayIndex = 3
                NewQueryGrid.Columns("AssDrop").DisplayIndex = 4
                NewQueryGrid.Columns("FormName").DisplayIndex = 5
                NewQueryGrid.Columns("PageNo").DisplayIndex = 6
                NewQueryGrid.Columns("Description").DisplayIndex = 7
                NewQueryGrid.Columns("Priority").DisplayIndex = 8
                NewQueryGrid.Columns("SiteClm").DisplayIndex = 9
                NewQueryGrid.Columns("TypeClm").DisplayIndex = 10
                NewQueryGrid.Columns("Person").DisplayIndex = 11
                NewQueryGrid.Columns("GroupClm").DisplayIndex = 12
                NewQueryGrid.Columns("PDF").DisplayIndex = 13
                NewQueryGrid.Columns("CopyQuery").DisplayIndex = 14
                NewQueryGrid.Columns("RespondClm").DisplayIndex = 15
                NewQueryGrid.Columns("StatusCmb").DisplayIndex = 16


                'Invisible
                NewQueryGrid.Columns("SiteCode").DisplayIndex = 17
                NewQueryGrid.Columns("RespondCode").DisplayIndex = 18
                NewQueryGrid.Columns("TypeCode").DisplayIndex = 19
                NewQueryGrid.Columns("Status").DisplayIndex = 20
                NewQueryGrid.Columns("Study").DisplayIndex = 21
                NewQueryGrid.Columns("QueryID").DisplayIndex = 22
                NewQueryGrid.Columns("CreatedBy").DisplayIndex = 23
                NewQueryGrid.Columns("CreateDate").DisplayIndex = 24
                NewQueryGrid.Columns("CreateTime").DisplayIndex = 25
                NewQueryGrid.Columns("ClosedDate").DisplayIndex = 26
                NewQueryGrid.Columns("ClosedTime").DisplayIndex = 27
                NewQueryGrid.Columns("ClosedBy").DisplayIndex = 28
                NewQueryGrid.Columns("ClosedByRole").DisplayIndex = 29
                NewQueryGrid.Columns("Bounced").DisplayIndex = 30

                NewQueryGrid.Columns("PDF").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                NewQueryGrid.Columns("TypeClm").DefaultCellStyle.WrapMode = DataGridViewTriState.True
                NewQueryGrid.Columns("RVLID").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                NewQueryGrid.Columns("Initials").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                NewQueryGrid.Columns("PageNo").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                NewQueryGrid.Columns("Priority").AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                NewQueryGrid.Columns("CopyQuery").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                NewQueryGrid.Columns("RespondClm").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                NewQueryGrid.Columns("StatusCmb").AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                NewQueryGrid.Columns("CreatedByRole").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                NewQueryGrid.Columns("SiteClm").AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                NewQueryGrid.Columns("TypeClm").AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                NewQueryGrid.Columns("Person").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                NewQueryGrid.Columns("GroupClm").AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                NewQueryGrid.Columns("AssDrop").AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                NewQueryGrid.Columns("VisitName").AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                NewQueryGrid.Columns("FormName").AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells

        End Select


    End Sub


    Private Sub NewQueryGrid_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles NewQueryGrid.CellDoubleClick

        If e.RowIndex < 0 Then Exit Sub
        If e.ColumnIndex = sender.columns("PDF").index Then
            If Role <> NewQueryGrid.Item("CreatedByRole", e.RowIndex).Value.ToString Then
                MsgBox("Cannot action this query due to role")
                Exit Sub
            End If
            If NewQueryGrid.Item("Status", e.RowIndex).Value <> "Responded" Then
                MsgBox("Query must be responded to link to PDF")
                Exit Sub
            End If
            'PDF Stuff
            Dim FilePath As String = ""
            Try
                FilePath = NewQueryGrid.Item("PDFLink", e.RowIndex).Value
            Catch ex As Exception
            End Try
            If FilePath = "" Then
                Dim fd As OpenFileDialog = New OpenFileDialog()

                fd.Title = "Open File Dialog"
                fd.InitialDirectory = "C:\"
                fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
                fd.FilterIndex = 2
                fd.RestoreDirectory = True

                If fd.ShowDialog() = DialogResult.OK Then
                    NewQueryGrid.Item("PDFLink", e.RowIndex).Value = fd.FileName
                End If
            Else
                If MsgBox("A file is already attached, do you want to replace it with another?", vbYesNo) = vbNo Then
                    Process.Start("explorer.exe", FilePath)
                Else
                    Dim fd As OpenFileDialog = New OpenFileDialog()

                    fd.Title = "Open File Dialog"
                    fd.InitialDirectory = "C:\"
                    fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
                    fd.FilterIndex = 2
                    fd.RestoreDirectory = True

                    If fd.ShowDialog() = DialogResult.OK Then
                        NewQueryGrid.Item("PDFLink", e.RowIndex).Value = fd.FileName
                    End If
                End If
            End If
        End If
        If e.ColumnIndex = sender.columns("CopyQuery").index Then
            If Role <> NewQueryGrid.Item("CreatedByRole", e.RowIndex).Value.ToString Then
                MsgBox("Cannot action this query due to role")
                Exit Sub
            End If
            If MsgBox("Do you want To copy this query To a New line?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                Dim NewRow As DataRow = Overclass.CurrentDataSet.Tables(0).NewRow
                NewRow.Item("VisitName") = NewQueryGrid.Item("VisitName", e.RowIndex).Value
                NewRow.Item("FormName") = NewQueryGrid.Item("FormName", e.RowIndex).Value
                NewRow.Item("PageNo") = NewQueryGrid.Item("PageNo", e.RowIndex).Value
                NewRow.Item("Description") = NewQueryGrid.Item("Description", e.RowIndex).Value
                NewRow.Item("Priority") = Trim(NewQueryGrid.Item("Priority", e.RowIndex).Value)
                NewRow.Item("Study") = Trim(NewQueryGrid.Item("Study", e.RowIndex).Value)
                NewRow.Item("RVLID") = NewQueryGrid.Item("RVLID", e.RowIndex).Value
                NewRow.Item("Initials") = NewQueryGrid.Item("Initials", e.RowIndex).Value
                NewRow.Item("AssCode") = NewQueryGrid.Item("AssCode", e.RowIndex).Value
                NewRow.Item("SiteCode") = Trim(NewQueryGrid.Item("SiteCode", e.RowIndex).Value)
                NewRow.Item("TypeCode") = Trim(NewQueryGrid.Item("TypeCode", e.RowIndex).Value)
                NewRow.Item("Person") = NewQueryGrid.Item("Person", e.RowIndex).Value
                NewRow.Item("RespondCode") = NewQueryGrid.Item("RespondCode", e.RowIndex).Value
                NewRow.Item("Status") = "Open"

                Overclass.CurrentDataSet.Tables(0).Rows.Add(NewRow)
                NewQueryGrid.CurrentCell = NewQueryGrid.Item("RVLID", NewQueryGrid.NewRowIndex)

            End If
        End If

        If e.ColumnIndex = sender.columns("StatusCmb").index Then

            If Role <> NewQueryGrid.Item("CreatedByRole", e.RowIndex).Value.ToString Then
                MsgBox("Cannot action this query due to role")
                Exit Sub
            End If

            If Me.NewQueryGrid.Item("Status", e.RowIndex).Value = "Closed" Then Exit Sub
            If IsDBNull(NewQueryGrid.Item("QueryID", e.RowIndex).Value) = True Then Exit Sub

            If MsgBox("Are you sure you want To close this query?", vbYesNo) = vbYes Then
                Me.NewQueryGrid.Item("ClosedDate", e.RowIndex).Value = Format(DateTime.Now, "dd-MMM-yyyy")
                Me.NewQueryGrid.Item("ClosedTime", e.RowIndex).Value = Format(DateTime.Now, "HH: mm")
                Me.NewQueryGrid.Item("ClosedBy", e.RowIndex).Value = Overclass.GetUserName
                Me.NewQueryGrid.Item("ClosedByRole", e.RowIndex).Value = Role
                Me.NewQueryGrid.Item("Status", e.RowIndex).Value = "Closed"
                BindingSource1.EndEdit()
            End If
        End If

        If e.ColumnIndex = sender.columns("RespondClm").index Then

            If Role <> NewQueryGrid.Item("CreatedByRole", e.RowIndex).Value.ToString Then
                MsgBox("Cannot action this query due to role")
                Exit Sub
            End If

            If NewQueryGrid.Item("Status", e.RowIndex).Value <> "Responded" Then Exit Sub

            Dim RespondText As String = vbNullString

            Dim QueryID As String
            QueryID = NewQueryGrid.Item("QueryID", e.RowIndex).Value

            If IsDBNull(NewQueryGrid.Item("QueryID", e.RowIndex).Value) = False Then
                Dim CSVString As String = Overclass.CreateCSVString(
                "Select format(Response_Timestamp,'dd-MMM-yyyy HH:mm') & ' (' & response_Person & ')  -  ' " &
                "& replace(response_text,',',';') FROM Response WHERE QueryID=" & QueryID)
                CSVString = Replace(CSVString, ",", vbNewLine & vbNewLine)
                If CSVString = "" Then CSVString = "No history found"
                RespondText = CSVString
            End If

            RespondText = RespondText & vbNewLine & "Please input response to query:"

            Dim Response = InputBox(RespondText, "Query Response")
            If Response = "" Then
                Exit Sub
            Else
                Dim SQL As String
                'NewQueryGrid.Item("Status", e.RowIndex).Value = "Open"
                Dim foundRows() As DataRow
                foundRows = Overclass.CurrentDataSet.Tables(0).Select("QueryID=" & QueryID)
                foundRows(0).Item("Bounced") = "True"
                foundRows(0).Item("Status") = "Open"
                foundRows(0).EndEdit()
                SQL = "INSERT INTO Response(QueryID,Response_Text,Response_Person) " &
                "VALUES (" & QueryID & ", '" & Response & "', '" & Overclass.GetUserName & "')"

                Dim Cmd As OleDb.OleDbCommand
                Cmd = New OleDb.OleDbCommand(SQL)
                Overclass.SetCommandConnection(Cmd)
                RespondCommands.Add(Cmd)
            End If
        End If

    End Sub

    Private Sub NewQueryGrid_DefaultValuesNeeded_1(sender As Object, e As DataGridViewRowEventArgs) Handles NewQueryGrid.DefaultValuesNeeded
        NewQueryGrid.Item("Status", e.Row.Index).Value = "Open"
    End Sub

    Private Sub NewQueryGrid_DataError(sender As Object, e As DataGridViewDataErrorEventArgs)

    End Sub

    Private Sub NewQueryGrid_RowPrePaint(sender As Object, e As DataGridViewRowPrePaintEventArgs) Handles NewQueryGrid.RowPrePaint

        Try

            If IsDBNull(NewQueryGrid.Item("QueryID", e.RowIndex).Value) Then
                If NewQueryGrid.Item("StatusCmb", e.RowIndex).Tag <> "Hyphen" Then
                    NewQueryGrid.Item("StatusCmb", e.RowIndex).Value = My.Resources.hyphen
                    NewQueryGrid.Item("StatusCmb", e.RowIndex).Tag = "Hyphen"
                End If
            End If

            If NewQueryGrid.Item("Status", e.RowIndex).Value = "Closed" Or Role <> NewQueryGrid.Item("CreatedByRole", e.RowIndex).Value Then
                If NewQueryGrid.Item("StatusCmb", e.RowIndex).Tag <> "Hyphen" Then
                    NewQueryGrid.Item("StatusCmb", e.RowIndex).Value = My.Resources.hyphen
                    NewQueryGrid.Item("StatusCmb", e.RowIndex).Tag = "Hyphen"
                End If
            End If

            If NewQueryGrid.Item("Status", e.RowIndex).Value <> "Responded" Or Role <> NewQueryGrid.Item("CreatedByRole", e.RowIndex).Value Then
                If NewQueryGrid.Item("RespondClm", e.RowIndex).Tag <> "Hyphen" Then
                    NewQueryGrid.Item("RespondClm", e.RowIndex).Value = My.Resources.hyphen
                    NewQueryGrid.Item("RespondClm", e.RowIndex).Tag = "Hyphen"
                End If
            End If

            If Role <> NewQueryGrid.Item("CreatedByRole", e.RowIndex).Value And NewQueryGrid.Item("CreatedByRole", e.RowIndex).Value <> "" Then
                If NewQueryGrid.Item("CopyQuery", e.RowIndex).Tag <> "Hyphen" Then
                    NewQueryGrid.Item("CopyQuery", e.RowIndex).Value = My.Resources.hyphen
                    NewQueryGrid.Item("CopyQuery", e.RowIndex).Tag = "Hyphen"
                    If NewQueryGrid.Rows(e.RowIndex).ReadOnly <> True Then NewQueryGrid.Rows(e.RowIndex).ReadOnly = True
                End If
            End If

            If Role <> NewQueryGrid.Item("CreatedByRole", e.RowIndex).Value Or NewQueryGrid.Item("Status", e.RowIndex).Value <> "Responded" Then
                If NewQueryGrid.Item("PDF", e.RowIndex).Tag <> "Hyphen" Then
                    NewQueryGrid.Item("PDF", e.RowIndex).Value = My.Resources.hyphen
                    NewQueryGrid.Item("PDF", e.RowIndex).Tag = "Hyphen"
                End If
            End If

                If NewQueryGrid.Item("CreatedByRole", e.RowIndex).ReadOnly <> True Then NewQueryGrid.Item("CreatedByRole", e.RowIndex).ReadOnly = True

        Catch ex As Exception

        End Try

    End Sub
End Class