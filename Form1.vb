Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Visible = False
        Me.Hide()

        Call StartUpCentral()

        If AccessLevel = 1 Then

            AdQry = New AddQuery
            AddControls(AdQry)
            AdQry.ShowDialog()

        ElseIf AccessLevel = 0 Then

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

        Select Case e.TabPage.Text

            Case "Add Queries"
                ctl = Me.DataGridView1
                SQLCode = "SELECT StudyCode, UploadDate, UploadPerson FROM Study ORDER BY UploadDate DESC"
                Overclass.CreateDataSet(SQLCode, Bind, ctl)

            Case "Query Codes"
                Specifics(DataGridView2)

            Case "Export"
                FilterCombo6.AllowBlanks = False
                FilterCombo6.SetAsExternalSource("StudyCode", "StudyCode", "SELECT StudyCode FROM Study", Overclass)

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
                SqlCode = "SELECT QueryID, Study, Person, " &
                "Priority, Initials & ' ' & RVLID AS Volunteer, VisitName, FormName, PageNo, Description, SiteCode, RespondCode " &
                "FROM Queries INNER JOIN Study ON Queries.Study=Study.StudyCode WHERE Hidden=False AND Status ='Open' ORDER BY Initials"
                Overclass.CreateDataSet(SqlCode, RespView.BindingSource1, RespView.StaffQueryGrid)

                With RespView.StaffQueryGrid
                    .ReadOnly = True
                    .Columns("QueryID").Visible = False
                    .Columns("Study").Visible = False
                    .Columns("SiteCode").Visible = False
                    .Columns("RespondCode").Visible = False
                    .Columns("Priority").Visible = False
                    .Columns("VisitName").HeaderText = "Study Visit"
                    .Columns("FormName").HeaderText = "Assessment/Procedure"
                    .Columns("PageNo").HeaderText = "Page No"
                    .Columns("Person").HeaderText = "Assigned"
                    Dim clm1 As New DataGridViewImageColumn
                    clm1.HeaderText = "History"
                    clm1.Name = "ViewClm"
                    clm1.ImageLayout = DataGridViewImageCellLayout.Zoom
                    clm1.Image = My.Resources.PreviousHistory
                    .Columns.Add(clm1)
                    Dim clm2 As New DataGridViewImageColumn
                    clm2.HeaderText = "Respond"
                    clm2.Name = "RespondClm"
                    clm2.ImageLayout = DataGridViewImageCellLayout.Zoom
                    clm2.Image = My.Resources.speech
                    .Columns.Add(clm2)
                    .Columns("RespondClm").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                    .Columns("ViewClm").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                    .Columns("PageNo").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                    .Columns("Person").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                    .Columns("Volunteer").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                End With

                With RespView
                    .FilterCombo3.AllowBlanks = False
                    .FilterCombo3.SetAsInternalSource("Study", "Study", Overclass)
                    .FilterCombo30.SetAsInternalSource("Priority", "Priority", Overclass)
                    .FilterCombo90.SetAsInternalSource("Volunteer", "Volunteer", Overclass)

                    .StaffQueryGrid.Columns("QueryID").Visible = False
                End With

                If CurrentFilter <> "" Then Overclass.CurrentDataSet.Tables(0).DefaultView.RowFilter = CurrentFilter
                RespView.FilterCombo1.Text = Filter1
                RespView.FilterCombo2.Text = Filter2
                RespView.FilterCombo3.Text = Filter3
                RespView.FilterCombo30.Text = Filter4
                RespView.FilterCombo90.Text = Filter5

            Case "NewQueryGrid2"

                If AdQry.CheckBox201.Checked = True Then
                    SqlCode = "SELECT Study, QueryID, SiteCode, TypeCode, RespondCode, RVLID, " &
                              "VisitName, FormName, Description, Status, Person FROM IncorrectQueries " &
                              "WHERE CreatedByRole='" & Role & "'" &
                              "ORDER BY RVLID ASC"
                Else
                    SqlCode = "SELECT Study, Queries.QueryID, SiteCode, TypeCode, RespondCode, RVLID, " &
                              "VisitName, FormName, Description, Status, Person FROM Queries INNER JOIN Study ON queries.Study=Study.StudyCode " &
                              "WHERE Hidden=false AND CreatedByRole='" & Role & "'" &
                              "ORDER BY RVLID ASC"
                End If



                Overclass.ResetCollection()
                Overclass.CreateDataSet(SqlCode, BindingSource1, ctl)

                AdQry.FilterCombo50.AllowBlanks = False
                AdQry.FilterCombo50.SetAsInternalSource("Study", "Study", Overclass)
                AdQry.FilterCombo40.SetAsInternalSource("RVLID", "RVLID", Overclass)
                AdQry.FilterCombo80.SetAsInternalSource("Status", "Status", Overclass)
                AdQry.FilterCombo70.SetAsInternalSource("VisitName", "VisitName", Overclass)
                AdQry.FilterCombo60.SetAsInternalSource("SiteCode", "SiteCode", Overclass)

                ctl.columns("QueryID").visible = False
                ctl.columns("TypeCode").visible = False
                ctl.columns("SiteCode").visible = False
                ctl.columns("RespondCode").visible = False
                ctl.columns("VisitName").readonly = True
                ctl.columns("FormName").readonly = True
                ctl.columns("Status").readonly = True
                ctl.columns("Description").readonly = True
                ctl.columns("RVLID").readonly = True
                ctl.columns("RVLID").headertext = "Subject ID"
                ctl.AllowUserToAddRows = False


                Dim cmb As TemplateDB.MyCmbColumn = Overclass.SetUpNewComboColumn("SELECT Code & ' - ' & Site AS Display, Code FROM SiteCode " &
                        "a inner join Study b ON a.ListID=b.CodeList " &
                        "WHERE CStr(StudyCode)=", AdQry.FilterCombo50, "Code", "Display", "SiteCode", "Site Code", AdQry.NewQueryGrid2, "clm1")

                Dim cmb2 As TemplateDB.MyCmbColumn = Overclass.SetUpNewComboColumn("SELECT Code & ' - ' & ErrorType AS Display, " &
                        "Code FROM TypeCode a inner join Study b ON a.ListID=b.CodeList " &
                        "WHERE Cstr(StudyCode)=", AdQry.FilterCombo50, "Code", "Display", "TypeCode", "Type Code", AdQry.NewQueryGrid2, "clm2")

                Dim cmb3 As TemplateDB.MyCmbColumn = Overclass.SetUpNewComboColumn("SELECT Code & ' - ' & Group AS Display, " &
                         "Code FROM GroupCode a inner join Study b ON a.ListID=b.CodeList " &
                         "WHERE CStr(StudyCode)=", AdQry.FilterCombo50, "Code", "Display", "RespondCode", "Respond Code", AdQry.NewQueryGrid2, "clm3")

                AdQry.NewQueryGrid2.Columns("FormName").HeaderText = "Form Name"
                Dim ctl3 As Object = AdQry.NewQueryGrid2
                ctl3.Columns("Person").maxinputlength = 3
                AdQry.NewQueryGrid2.Columns("Person").DisplayIndex = AdQry.NewQueryGrid2.Columns.Count - 2


            Case "NewQueryGrid"

                AdQry.NewQueryGrid.Columns.Clear()

                SqlCode = "SELECT Status, Study, QueryID, CreatedBy, FieldName, CreateDate, CreateTime, " &
                    "CreatedByRole, ClosedDate, ClosedTime, ClosedBy, ClosedByRole, RVLID, Initials, " &
                    "VisitName, FormName, PageNo, Description, Priority " &
                    "FROM Queries INNER JOIN Study On queries.Study=Study.StudyCode " &
                    "WHERE Hidden=false AND QueryID Like 'MANUAL-%' " &
                    "AND CreatedByRole='" & Role & "' " &
                    "ORDER BY Status DESC, RVLID ASC"

                Overclass.CreateDataSet(SqlCode, AdQry.BindingSource1, AdQry.NewQueryGrid)

                AdQry.FilterCombo30.AllowBlanks = False
                AdQry.FilterCombo30.SetAsExternalSource("Study", "Study", "SELECT StudyCode As Study FROM Study " &
                "WHERE Hidden=False ORDER BY StudyCode ASC", Overclass)
                AdQry.FilterCombo30.SetDGVDefault(ctl, "Study")

                AdQry.FilterCombo90.SetAsInternalSource("Initials", "Initials", Overclass)
                AdQry.FilterCombo100.SetAsInternalSource("Status", "Status", Overclass)

                AdQry.FilterCombo20.SetAsInternalSource("RVLID", "RVLID", Overclass)
                AdQry.FilterCombo10.SetAsInternalSource("VisitName", "VisitName", Overclass)

                AdQry.NewQueryGrid.Columns("QueryID").Visible = False
                AdQry.NewQueryGrid.Columns("Study").Visible = False
                AdQry.NewQueryGrid.Columns("FieldName").Visible = False
                AdQry.NewQueryGrid.Columns("CreateDate").Visible = False
                AdQry.NewQueryGrid.Columns("CreateTime").Visible = False
                AdQry.NewQueryGrid.Columns("CreatedByRole").Visible = False
                AdQry.NewQueryGrid.Columns("ClosedDate").Visible = False
                AdQry.NewQueryGrid.Columns("ClosedTime").Visible = False
                AdQry.NewQueryGrid.Columns("ClosedBy").Visible = False
                AdQry.NewQueryGrid.Columns("ClosedByRole").Visible = False
                AdQry.NewQueryGrid.Columns("Study").Visible = False
                AdQry.NewQueryGrid.Columns("Priority").Visible = False
                AdQry.NewQueryGrid.Columns("Status").Visible = False

                AdQry.NewQueryGrid.Columns("Study").ReadOnly = False
                AdQry.NewQueryGrid.Columns("Status").ReadOnly = True
                AdQry.NewQueryGrid.Columns("CreatedBy").ReadOnly = True

                AdQry.NewQueryGrid.Columns("RVLID").HeaderText = "Subject ID"
                AdQry.NewQueryGrid.Columns("VisitName").HeaderText = "Study Visit"
                AdQry.NewQueryGrid.Columns("FormName").HeaderText = "Assessment/Procedure"
                AdQry.NewQueryGrid.Columns("PageNo").HeaderText = "Page No."
                AdQry.NewQueryGrid.Columns("Description").HeaderText = "Query Description"
                AdQry.NewQueryGrid.Columns("CreatedBy").HeaderText = "Created By"

                AdQry.NewQueryGrid.Columns("CreatedBy").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                AdQry.NewQueryGrid.Columns("RVLID").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                AdQry.NewQueryGrid.Columns("Initials").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                AdQry.NewQueryGrid.Columns("PageNo").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader

                Dim clm As New DataGridViewComboBoxColumn
                clm.HeaderText = "Priority"
                clm.Items.Add("1 - Data Entry")
                clm.Items.Add("2 - Non Data Entry")
                AdQry.NewQueryGrid.Columns.Add(clm)
                clm.DataPropertyName = "Priority"
                clm.Name = "PriorityClm"

                AdQry.NewQueryGrid.Columns("PriorityClm").AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells

                Dim cmb2 As New DataGridViewImageColumn
                cmb2.HeaderText = "Copy"
                cmb2.Image = My.Resources.copy
                cmb2.ImageLayout = DataGridViewImageCellLayout.Zoom
                cmb2.Name = "CopyQuery"

                AdQry.NewQueryGrid.Columns.Add(cmb2)
                AdQry.NewQueryGrid.Columns("CopyQuery").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader

                Dim clm1 As New DataGridViewImageColumn
                clm1.HeaderText = "History"
                clm1.Name = "ViewClm"
                clm1.ImageLayout = DataGridViewImageCellLayout.Zoom
                clm1.Image = My.Resources.PreviousHistory
                AdQry.NewQueryGrid.Columns.Add(clm1)
                AdQry.NewQueryGrid.Columns("ViewClm").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader

                Dim clm2 As New DataGridViewImageColumn
                clm2.HeaderText = "Respond"
                clm2.Name = "RespondClm"
                clm2.ImageLayout = DataGridViewImageCellLayout.Zoom
                clm2.Image = My.Resources.speech
                AdQry.NewQueryGrid.Columns.Add(clm2)
                AdQry.NewQueryGrid.Columns("RespondClm").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader

                Dim cmb As New DataGridViewImageColumn
                cmb.HeaderText = "Close"
                cmb.Image = My.Resources.TICK
                cmb.ImageLayout = DataGridViewImageCellLayout.Zoom
                cmb.Name = "StatusCmb"

                AdQry.NewQueryGrid.Columns.Add(cmb)
                AdQry.NewQueryGrid.Columns("StatusCmb").AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells


            Case "DataGridView1"
                ctl.columns(0).headertext = "Study"
                ctl.columns(1).headertext = "Last Update"
                ctl.columns(2).headertext = "Upload Person"
                ctl.columns(1).DefaultCellStyle.Format = "dd-MMM-yyyy - HH: mm"
                ctl.enabled = False
                ctl.AllowUserToAddRows = False

            Case "DataGridView2"

                If Me.CheckBox1.Checked = True Then
                    SqlCode = "SELECT Study, QueryID, SiteCode, TypeCode, RespondCode, RVLID, " &
                              "VisitName, FormName, Description, Status, Person FROM IncorrectQueries " &
                              "ORDER BY RVLID ASC"
                Else
                    SqlCode = "SELECT Study, Queries.QueryID, SiteCode, TypeCode, RespondCode, RVLID, " &
                              "VisitName, FormName, Description, Status, Person FROM Queries INNER JOIN Study ON queries.Study=Study.StudyCode " &
                              "WHERE Hidden=false ORDER BY RVLID ASC"
                End If



                Overclass.ResetCollection()
                Overclass.CreateDataSet(SqlCode, BindingSource1, ctl)

                FilterCombo2.AllowBlanks = False
                FilterCombo2.SetAsInternalSource("Study", "Study", Overclass)
                FilterCombo1.SetAsInternalSource("RVLID", "RVLID", Overclass)
                FilterCombo3.SetAsInternalSource("Status", "Status", Overclass)
                FilterCombo4.SetAsInternalSource("VisitName", "VisitName", Overclass)
                FilterCombo5.SetAsInternalSource("SiteCode", "SiteCode", Overclass)

                ctl.columns("QueryID").visible = False
                ctl.columns("TypeCode").visible = False
                ctl.columns("SiteCode").visible = False
                ctl.columns("RespondCode").visible = False
                ctl.columns("VisitName").readonly = True
                ctl.columns("FormName").readonly = True
                ctl.columns("Status").readonly = True
                ctl.columns("Description").readonly = True
                ctl.columns("RVLID").readonly = True
                ctl.columns("RVLID").headertext = "Subject ID"
                ctl.columns("Study").visible = False
                ctl.AllowUserToAddRows = False

                Dim cmb As TemplateDB.MyCmbColumn = Overclass.SetUpNewComboColumn("SELECT Code & ' - ' & Site AS Display, Code FROM SiteCode " &
                       "a inner join Study b ON a.ListID=b.CodeList " &
                       "WHERE CStr(StudyCode)=", FilterCombo2, "Code", "Display", "SiteCode", "Site Code", DataGridView2, "clm1")

                Dim cmb2 As TemplateDB.MyCmbColumn = Overclass.SetUpNewComboColumn("SELECT Code & ' - ' & ErrorType AS Display, " &
                       "Code FROM TypeCode a inner join Study b ON a.ListID=b.CodeList " &
                       "WHERE Cstr(StudyCode)=", FilterCombo2, "Code", "Display", "TypeCode", "Type Code", DataGridView2, "clm2")

                Dim cmb3 As TemplateDB.MyCmbColumn = Overclass.SetUpNewComboColumn("SELECT Code & ' - ' & Group AS Display, " &
                        "Code FROM GroupCode a inner join Study b ON a.ListID=b.CodeList " &
                        "WHERE CStr(StudyCode)=", FilterCombo2, "Code", "Display", "RespondCode", "Respond Code", DataGridView2, "clm3")

                ctl.columns("Person").displayindex = 13
                ctl.columns("FormName").HeaderText = "Form Name"
                ctl.columns("VisitName").HeaderText = "Visit Name"
                ctl.Columns("Person").maxinputlength = 3

                cmb.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                cmb3.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                DataGridView2.Columns("Person").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                DataGridView2.Columns("RVLID").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                DataGridView2.Columns("Status").Visible = False



        End Select


    End Sub

    Private Sub CheckBox1_Click(sender As Object, e As EventArgs) Handles CheckBox1.Click

        If Overclass.UnloadData() = True Then
            RemoveHandler CheckBox1.Click, AddressOf CheckBox1_Click
            CheckBox1.Checked = Not CheckBox1.Checked
            AddHandler CheckBox1.Click, AddressOf CheckBox1_Click
        End If

        Dim TempStudy As String = FilterCombo2.SelectedValue
        Call Specifics(Me.DataGridView2)
        If TempStudy IsNot Nothing Then FilterCombo2.SelectedValue = TempStudy
        FilterCombo1.SelectedValue = ""
        FilterCombo3.SelectedValue = ""
        FilterCombo4.SelectedValue = ""
        FilterCombo5.SelectedValue = ""

    End Sub

End Class