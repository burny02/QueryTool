Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Visible = False
        Me.Hide()

        Call StartUpCentral()

        If AccessLevel = 1 Then

            AdQry = New AddQuery
            AddControls(AdQry)
            AdQry.ShowDialog()

        Else

            Me.Visible = True
            Me.WindowState = FormWindowState.Maximized

        End If


        Try
            Me.Label2.Text = "Query Tool " & vbNewLine & "Developed by David Burnside" & vbNewLine & "Version: " & System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
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
                SQLCode = "SELECT DisplayName, UploadDate, UploadPerson FROM Study ORDER BY UploadDate DESC"
                Overclass.CreateDataSet(SQLCode, Bind, ctl)

            Case "Query Codes"
                Specifics(DataGridView2)

            Case "Export"
                FilterCombo6.AllowBlanks = False
                FilterCombo6.SetAsExternalSource("StudyCode", "DisplayName", "SELECT StudyCode, DisplayName FROM Study", Overclass)

            Case "Reports"
                Me.DateTimePicker2.Value = Date.Now

        End Select


        Call Specifics(ctl)

    End Sub

    Public Sub Specifics(ctl As Object)

        If IsNothing(ctl) Then Exit Sub

        Dim SqlCode As String = vbNullString

        Select Case ctl.name


            Case "NewQueryGrid2"

                If AdQry.CheckBox201.Checked = True Then
                    SqlCode = "SELECT Study, DisplayName, QueryID, SiteCode, TypeCode, RespondCode, RVLID, " &
                              "VisitName, FormName, Description, Status, Person FROM IncorrectQueries " &
                              "WHERE CreatedByRole='" & Role & "'" &
                              "ORDER BY RVLID ASC"
                Else
                    SqlCode = "SELECT Study, DisplayName, Queries.QueryID, SiteCode, TypeCode, RespondCode, RVLID, " &
                              "VisitName, FormName, Description, Status, Person FROM (Queries " &
                              "INNER JOIN QueryCodes ON QueryCodes.QueryID = Queries.QueryID) " &
                              "INNER JOIN Study ON Queries.Study=Study.StudyCode " &
                              "WHERE CreatedByRole='" & Role & "'" &
                              "ORDER BY RVLID ASC"
                End If



                Overclass.ResetCollection()
                Overclass.CreateDataSet(SqlCode, BindingSource1, ctl)

                AdQry.FilterCombo50.AllowBlanks = False
                AdQry.FilterCombo50.SetAsInternalSource("Study", "DisplayName", Overclass)
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
                ctl.columns("Study").visible = False
                ctl.columns("DisplayName").headertext = "Study"
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

                SqlCode = "SELECT Status, DisplayName, QueryID, CreatedBy, FieldName, CreateDate, CreateTime, " &
                    "CreatedByRole, ClosedDate, ClosedTime, ClosedBy, ClosedByRole, RVLID, Initials, " &
                    "VisitName, FormName, PageNo, Description, Priority, ResolvedBy, ResolvedDate, Study " &
                    "FROM Queries INNER JOIN Study ON Queries.Study=Study.StudyCode " &
                    "WHERE QueryID Like 'MANUAL-%' " &
                    "AND CreatedByRole='" & Role & "' " &
                    "ORDER BY Status DESC, RVLID ASC"

                Overclass.CreateDataSet(SqlCode, AdQry.BindingSource1, AdQry.NewQueryGrid)

                AdQry.FilterCombo30.AllowBlanks = False
                AdQry.FilterCombo30.SetAsExternalSource("Study", "DisplayName", "SELECT StudyCode As Study, DisplayName FROM Study ORDER BY DisplayName ASC", Overclass)
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
                AdQry.NewQueryGrid.Columns("DisplayName").Visible = False
                AdQry.NewQueryGrid.Columns("Priority").Visible = False
                AdQry.NewQueryGrid.Columns("Status").Visible = False

                AdQry.NewQueryGrid.Columns("DisplayName").ReadOnly = False
                AdQry.NewQueryGrid.Columns("Status").ReadOnly = True
                AdQry.NewQueryGrid.Columns("CreatedBy").ReadOnly = True

                AdQry.NewQueryGrid.Columns("RVLID").HeaderText = "Subject ID"
                AdQry.NewQueryGrid.Columns("VisitName").HeaderText = "Study Visit"
                AdQry.NewQueryGrid.Columns("FormName").HeaderText = "Assessment/Procedure"
                AdQry.NewQueryGrid.Columns("PageNo").HeaderText = "Page No."
                AdQry.NewQueryGrid.Columns("Description").HeaderText = "Query Description"
                AdQry.NewQueryGrid.Columns("CreatedBy").HeaderText = "Created By"
                AdQry.NewQueryGrid.Columns("ResolvedBy").HeaderText = "Resolved By"
                AdQry.NewQueryGrid.Columns("ResolvedDate").HeaderText = "Resolved Date"


                AdQry.NewQueryGrid.Columns("ResolvedDate").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                AdQry.NewQueryGrid.Columns("CreatedBy").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                AdQry.NewQueryGrid.Columns("ResolvedBy").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
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

                Dim cmb As New DataGridViewComboBoxColumn
                cmb.HeaderText = "Status"
                cmb.Items.Add("Open")
                cmb.Items.Add("Closed")
                cmb.DataPropertyName = "Status"
                cmb.Name = "StatusCmb"

                AdQry.NewQueryGrid.Columns.Add(cmb)
                AdQry.NewQueryGrid.Columns("StatusCmb").AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells

                Dim cmb2 As New DataGridViewImageColumn
                cmb2.HeaderText = "Copy"
                cmb2.Image = My.Resources.copy
                cmb2.ImageLayout = DataGridViewImageCellLayout.Zoom
                cmb2.Name = "CopyQuery"

                AdQry.NewQueryGrid.Columns.Add(cmb2)
                AdQry.NewQueryGrid.Columns("CopyQuery").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader


            Case "DataGridView1"
                ctl.columns(0).headertext = "Study"
                ctl.columns(1).headertext = "Last Update"
                ctl.columns(2).headertext = "Upload Person"
                ctl.columns(1).DefaultCellStyle.Format = "dd-MMM-yyyy - HH: mm"
                ctl.enabled = False
                ctl.AllowUserToAddRows = False

            Case "DataGridView2"

                If Me.CheckBox1.Checked = True Then
                    SqlCode = "SELECT Study, DisplayName, QueryID, SiteCode, TypeCode, RespondCode, RVLID, " &
                              "VisitName, FormName, Description, Status, Person FROM IncorrectQueries " &
                              "ORDER BY RVLID ASC"
                Else
                    SqlCode = "SELECT Study, DisplayName, Queries.QueryID, SiteCode, TypeCode, RespondCode, RVLID, " &
                              "VisitName, FormName, Description, Status, Person FROM (Queries " &
                              "INNER JOIN QueryCodes ON QueryCodes.QueryID = Queries.QueryID) " &
                              "INNER JOIN Study ON Queries.Study=Study.StudyCode " &
                              "ORDER BY RVLID ASC"
                End If



                Overclass.ResetCollection()
                Overclass.CreateDataSet(SqlCode, BindingSource1, ctl)

                FilterCombo2.AllowBlanks = False
                FilterCombo2.SetAsInternalSource("Study", "DisplayName", Overclass)
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
                ctl.columns("DisplayName").visible = False
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