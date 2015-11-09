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

        If overclass.UnloadData() = True Then
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
                StartCombo(Me.ComboBox1)
                Specifics(Me.DataGridView2)
                StartCombo(Me.ComboBox8)
                StartCombo(Me.ComboBox7)
                StartCombo(Me.ComboBox6)
                StartCombo(Me.ComboBox5)
                StartCombo(Me.ComboBox9)

            Case "Export"
                StartCombo(Me.ComboBox3)

            Case "Reports"
                Me.DateTimePicker2.Value = Date.Now

        End Select


        Call Specifics(ctl)

    End Sub

    Public Sub Specifics(ctl As Object)

        If IsNothing(ctl) Then Exit Sub

        Dim SqlCode As String = vbNullString

        Select Case ctl.name

            Case "NewQueryGrid3"

                AdQry.NewQueryGrid3.Columns.Clear()

                If IsNothing(AdQry.ComboBox108.SelectedValue) Then Exit Sub


                SqlCode = "SELECT a.QueryID, CreatedBy, RVLID, VisitName, FormName, SiteCode, TypeCode, Person, RespondCode, " & _
                                "Description FROM QueryCodes as a INNER JOIN Queries as b ON a.QueryID=b.QueryID " & _
                                "WHERE Study='" & AdQry.ComboBox108.SelectedValue.ToString & "' AND a.QueryID Like 'MANUAL-%' " & _
                                "AND CreatedByRole='" & Role & "'" & _
                                "ORDER BY RVLID ASC"


                Overclass.CreateDataSet(SqlCode, AdQry.BindingSource1, AdQry.NewQueryGrid3)

                AdQry.NewQueryGrid3.Columns("QueryID").Visible = False
                AdQry.NewQueryGrid3.Columns("SiteCode").Visible = False
                AdQry.NewQueryGrid3.Columns("TypeCode").Visible = False
                AdQry.NewQueryGrid3.Columns("RespondCode").Visible = False
                AdQry.NewQueryGrid3.Columns("RVLID").ReadOnly = True
                AdQry.NewQueryGrid3.Columns("FormName").ReadOnly = True
                AdQry.NewQueryGrid3.Columns("Description").ReadOnly = True
                AdQry.NewQueryGrid3.Columns("VisitName").ReadOnly = True
                AdQry.NewQueryGrid3.Columns("FormName").ReadOnly = True
                AdQry.NewQueryGrid3.Columns("CreatedBy").ReadOnly = True

                Dim cmb As New DataGridViewComboBoxColumn()
                cmb.DataSource = Overclass.TempDataTable("SELECT Code & ' - ' & Site AS Display, Code FROM SiteCode a inner join Study b ON a.ListID=b.CodeList " & _
                                             "WHERE StudyCode='" & AdQry.ComboBox108.SelectedValue.ToString & "' ORDER BY Code ASC")
                cmb.DataPropertyName = Overclass.CurrentDataSet.Tables(0).Columns("SiteCode").ToString
                cmb.ValueMember = "Code"
                cmb.DisplayMember = "Display"
                AdQry.NewQueryGrid3.Columns.Add(cmb)
                Dim cmb2 As New DataGridViewComboBoxColumn()
                cmb2.DataSource = Overclass.TempDataTable("SELECT Code & ' - ' & ErrorType AS Display, Code FROM TypeCode a inner join Study b ON a.ListID=b.CodeList " & _
                                             "WHERE StudyCode='" & AdQry.ComboBox108.SelectedValue.ToString & "' ORDER BY Code ASC")
                cmb2.DataPropertyName = Overclass.CurrentDataSet.Tables(0).Columns("TypeCode").ToString
                cmb2.ValueMember = "Code"
                cmb2.DisplayMember = "Display"
                AdQry.NewQueryGrid3.Columns.Add(cmb2)
                Dim cmb3 As New DataGridViewComboBoxColumn()
                cmb3.DataSource = Overclass.TempDataTable("SELECT Code & ' - ' & Group AS Display, Code FROM GroupCode a inner join Study b ON a.ListID=b.CodeList " & _
                                             "WHERE StudyCode='" & AdQry.ComboBox108.SelectedValue.ToString & "' ORDER BY Code ASC")
                cmb3.DataPropertyName = Overclass.CurrentDataSet.Tables(0).Columns("RespondCode").ToString
                cmb3.ValueMember = "Code"
                cmb3.DisplayMember = "Display"
                AdQry.NewQueryGrid3.Columns.Add(cmb3)
                cmb3.HeaderText = "Respond Code"
                cmb2.HeaderText = "Type Code"
                cmb.HeaderText = "Site Code"
                AdQry.NewQueryGrid3.Columns("FormName").HeaderText = "Form Name"
                Dim ctl2 As Object = AdQry.NewQueryGrid3
                ctl2.Columns("Person").maxinputlength = 3
                AdQry.NewQueryGrid3.Columns("Person").DisplayIndex = AdQry.NewQueryGrid3.Columns.Count - 2



            Case "NewQueryGrid2"

                AdQry.NewQueryGrid2.Columns.Clear()

                If IsNothing(AdQry.ComboBox107.SelectedValue) Then Exit Sub

                Dim AllowedSite As String = Overclass.CreateCSVString("SELECT Code FROM SiteCODE a INNER JOIN Study b ON a.ListID=b.CodeList " & _
                                                            "WHERE StudyCode='" & AdQry.ComboBox107.SelectedValue.ToString & "'")
                Dim AllowedResponse As String = Overclass.CreateCSVString("SELECT Code FROM GroupCode a INNER JOIN Study b ON a.ListID=b.CodeList " & _
                                                            "WHERE StudyCode='" & AdQry.ComboBox107.SelectedValue.ToString & "'")
                Dim AllowedType As String = Overclass.CreateCSVString("SELECT Code FROM TypeCode a INNER JOIN Study b ON a.ListID=b.CodeList " & _
                                                            "WHERE StudyCode='" & AdQry.ComboBox107.SelectedValue.ToString & "'")

                SqlCode = "SELECT a.QueryID, CreatedBy, RVLID, VisitName, FormName, SiteCode, TypeCode, Person, RespondCode, " & _
                                "Description FROM QueryCodes as a INNER JOIN Queries as b ON a.QueryID=b.QueryID " & _
                                "WHERE Study='" & AdQry.ComboBox107.SelectedValue.ToString & "'" & _
                                " AND a.QueryID Like 'MANUAL-%' " & _
                                "AND CreatedByRole='" & Role & "'" & _
                                "AND (instr('" & AllowedSite & "',SiteCode)=0" & _
                                " OR instr('" & AllowedResponse & "',RespondCode)=0" & _
                                " OR instr('" & AllowedType & "',TypeCode)=0" & _
                                " OR SiteCode=''" & _
                                " OR RespondCode=''" & _
                                " OR Person=''" & _
                                " OR Person NOT Like '[a-z][a-z-][a-z]'" & _
                                " OR isnull(Person)" & _
                                " OR isnull(SiteCode)" & _
                                " OR isnull(RespondCode)" & _
                                " OR isnull(TypeCode)" & _
                                " OR len(Person)<>3" & _
                                " OR TypeCode='')" & _
                                " ORDER BY RVLID ASC"

                Overclass.CreateDataSet(SqlCode, AdQry.BindingSource1, AdQry.NewQueryGrid2)
                AdQry.NewQueryGrid2.Columns("QueryID").Visible = False
                AdQry.NewQueryGrid2.Columns("SiteCode").Visible = False
                AdQry.NewQueryGrid2.Columns("TypeCode").Visible = False
                AdQry.NewQueryGrid2.Columns("RespondCode").Visible = False
                AdQry.NewQueryGrid2.Columns("RVLID").ReadOnly = True
                AdQry.NewQueryGrid2.Columns("FormName").ReadOnly = True
                AdQry.NewQueryGrid2.Columns("Description").ReadOnly = True
                AdQry.NewQueryGrid2.Columns("VisitName").ReadOnly = True
                AdQry.NewQueryGrid2.Columns("FormName").ReadOnly = True
                AdQry.NewQueryGrid2.Columns("CreatedBy").ReadOnly = True

                AdQry.NewQueryGrid2.AllowUserToAddRows = False
                Dim cmb As New DataGridViewComboBoxColumn()
                cmb.DataSource = Overclass.TempDataTable("SELECT Code & ' - ' & Site AS Display, Code FROM SiteCode a inner join Study b ON a.ListID=b.CodeList " & _
                                             "WHERE StudyCode='" & AdQry.ComboBox107.SelectedValue.ToString & "' ORDER BY Code ASC")
                cmb.DataPropertyName = Overclass.CurrentDataSet.Tables(0).Columns("SiteCode").ToString
                cmb.ValueMember = "Code"
                cmb.DisplayMember = "Display"
                AdQry.NewQueryGrid2.Columns.Add(cmb)
                Dim cmb2 As New DataGridViewComboBoxColumn()
                cmb2.DataSource = Overclass.TempDataTable("SELECT Code & ' - ' & ErrorType AS Display, Code FROM TypeCode a inner join Study b ON a.ListID=b.CodeList " & _
                                             "WHERE StudyCode='" & AdQry.ComboBox107.SelectedValue.ToString & "' ORDER BY Code ASC")
                cmb2.DataPropertyName = Overclass.CurrentDataSet.Tables(0).Columns("TypeCode").ToString
                cmb2.ValueMember = "Code"
                cmb2.DisplayMember = "Display"
                AdQry.NewQueryGrid2.Columns.Add(cmb2)
                Dim cmb3 As New DataGridViewComboBoxColumn()
                cmb3.DataSource = Overclass.TempDataTable("SELECT Code & ' - ' & Group AS Display, Code FROM GroupCode a inner join Study b ON a.ListID=b.CodeList " & _
                                             "WHERE StudyCode='" & AdQry.ComboBox107.SelectedValue.ToString & "' ORDER BY Code ASC")
                cmb3.DataPropertyName = Overclass.CurrentDataSet.Tables(0).Columns("RespondCode").ToString
                cmb3.ValueMember = "Code"
                cmb3.DisplayMember = "Display"
                AdQry.NewQueryGrid2.Columns.Add(cmb3)
                cmb3.HeaderText = "Respond Code"
                cmb2.HeaderText = "Type Code"
                cmb.HeaderText = "Site Code"
                AdQry.NewQueryGrid2.Columns("FormName").HeaderText = "Form Name"
                Dim ctl3 As Object = AdQry.NewQueryGrid2
                ctl3.Columns("Person").maxinputlength = 3
                AdQry.NewQueryGrid2.Columns("Person").DisplayIndex = AdQry.NewQueryGrid2.Columns.Count - 2


            Case "NewQueryGrid"

                AdQry.NewQueryGrid.Columns.Clear()

                If IsNothing(AdQry.ComboBox101.SelectedValue) Then Exit Sub

                Dim IDCrit As String = "'%'"
                Dim InitCrit As String = "'%'"
                Dim VisitCrit As String = "'%'"

                If AdQry.ComboBox102.SelectedValue <> "" Then IDCrit = "'" & AdQry.ComboBox102.SelectedValue & "'"
                If AdQry.ComboBox103.SelectedValue <> "" Then InitCrit = "'" & AdQry.ComboBox103.SelectedValue & "'"
                If AdQry.ComboBox105.SelectedValue <> "" Then VisitCrit = "'" & AdQry.ComboBox105.SelectedValue & "'"

                SqlCode = "SELECT QueryID, CreatedBy, Status, Study, FieldName, CreateDate, CreateTime, CreatedByRole, ClosedDate, ClosedTime, " & _
                    "ClosedBy, ClosedByRole, RVLID, Initials, VisitName, FormName, PageNo, Description " & _
                    "FROM Queries " & _
                    "WHERE Study='" & AdQry.ComboBox101.SelectedValue.ToString & "' " & _
                    "AND QueryID Like 'MANUAL-%' " & _
                    "AND CreatedByRole='" & Role & "' " & _
                    "AND Status='Open' " & _
                    " AND RVLID LIKE" & IDCrit & _
                    " AND Initials LIKE " & InitCrit & _
                    " AND VisitName LIKE " & VisitCrit & _
                    "ORDER BY RVLID ASC"

                Overclass.CreateDataSet(SqlCode, AdQry.BindingSource1, AdQry.NewQueryGrid)

                Dim cmb As New DataGridViewImageColumn
                cmb.HeaderText = "Close Query"
                cmb.Image = My.Resources.TICK
                cmb.ImageLayout = DataGridViewImageCellLayout.Zoom
                cmb.Name = "CloseQuery"


                AdQry.NewQueryGrid.Columns.Add(cmb)

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

                AdQry.NewQueryGrid.Columns("Status").ReadOnly = True
                AdQry.NewQueryGrid.Columns("CreatedBy").ReadOnly = True

                AdQry.NewQueryGrid.Columns("VisitName").HeaderText = "Study Visit"
                AdQry.NewQueryGrid.Columns("FormName").HeaderText = "Assessment/Procedure"
                AdQry.NewQueryGrid.Columns("PageNo").HeaderText = "Page No."
                AdQry.NewQueryGrid.Columns("Description").HeaderText = "Query Description"
                AdQry.NewQueryGrid.Columns("CreatedBy").HeaderText = "Created By"

            Case "DataGridView1"
                ctl.columns(0).headertext = "Study"
                ctl.columns(1).headertext = "Last Update"
                ctl.columns(2).headertext = "Upload Person"
                ctl.columns(1).DefaultCellStyle.Format = "dd-MMM-yyyy - HH:mm"
                ctl.enabled = False
                ctl.AllowUserToAddRows = False

            Case "DataGridView2"

                If IsNothing(Me.ComboBox1.SelectedValue) Then Exit Sub

                Dim IDCrit As String = "'%'"
                Dim InitCrit As String = "'%'"
                Dim StatusCrit As String = "'%'"
                Dim VisitCrit As String = "'%'"
                Dim SiteCrit As String = "'%'"

                If Me.ComboBox8.SelectedValue <> "" Then IDCrit = "'" & Me.ComboBox8.SelectedValue & "'"
                If Me.ComboBox7.SelectedValue <> "" Then InitCrit = "'" & Me.ComboBox7.SelectedValue & "'"
                If Me.ComboBox6.SelectedValue <> "" Then StatusCrit = "'" & Me.ComboBox6.SelectedValue & "'"
                If Me.ComboBox5.SelectedValue <> "" Then VisitCrit = "'" & Me.ComboBox5.SelectedValue & "'"
                If Me.ComboBox9.SelectedValue <> "" Then SiteCrit = "'" & Me.ComboBox9.SelectedValue & "'"

                If Me.CheckBox1.Checked = True Then

                    Dim AllowedSite As String = Overclass.CreateCSVString("SELECT Code FROM SiteCODE a INNER JOIN Study b ON a.ListID=b.CodeList " & _
                                                            "WHERE StudyCode='" & Me.ComboBox1.SelectedValue.ToString & "'")
                    Dim AllowedResponse As String = Overclass.CreateCSVString("SELECT Code FROM GroupCode a INNER JOIN Study b ON a.ListID=b.CodeList " & _
                                                                "WHERE StudyCode='" & Me.ComboBox1.SelectedValue.ToString & "'")
                    Dim AllowedType As String = Overclass.CreateCSVString("SELECT Code FROM TypeCode a INNER JOIN Study b ON a.ListID=b.CodeList " & _
                                                                "WHERE StudyCode='" & Me.ComboBox1.SelectedValue.ToString & "'")

                    SqlCode = "SELECT a.QueryID, SiteCode, TypeCode, Person, RespondCode, RVLID, " & _
                                    "VisitName, FormName, Description, Status FROM QueryCodes as a INNER JOIN Queries as b ON a.QueryID=b.QueryID " & _
                                    "WHERE Study='" & Me.ComboBox1.SelectedValue.ToString & "'" & _
                                    " AND a.QueryID LIKE" & IDCrit & _
                                    " AND Initials LIKE " & InitCrit & _
                                    " AND Status LIKE " & StatusCrit & _
                                    " AND VisitName LIKE " & VisitCrit & _
                                    " AND SiteCode LIKE " & SiteCrit & _
                                    "AND (instr('" & AllowedSite & "',SiteCode)=0" & _
                                    " OR instr('" & AllowedResponse & "',RespondCode)=0" & _
                                    " OR instr('" & AllowedType & "',TypeCode)=0" & _
                                    " OR SiteCode=''" & _
                                    " OR RespondCode=''" & _
                                    " OR Person=''" & _
                                    " OR Person NOT Like '[a-z][a-z-][a-z]'" & _
                                    " OR isnull(Person)" & _
                                    " OR isnull(SiteCode)" & _
                                    " OR isnull(RespondCode)" & _
                                    " OR isnull(TypeCode)" & _
                                    " OR len(Person)<>3" & _
                                    " OR TypeCode='')" & _
                                    " ORDER BY RVLID ASC"


                Else

                    SqlCode = "SELECT a.QueryID, SiteCode, TypeCode, Person, RespondCode, RVLID, " & _
                                "VisitName, FormName, Description, Status FROM QueryCodes as a INNER JOIN Queries as b ON a.QueryID=b.QueryID " & _
                                "WHERE Study='" & Me.ComboBox1.SelectedValue.ToString & "'" & _
                                " AND a.QueryID LIKE" & IDCrit & _
                                    " AND Initials LIKE " & InitCrit & _
                                    " AND Status LIKE " & StatusCrit & _
                                    " AND VisitName LIKE " & VisitCrit & _
                                    " AND SiteCode LIKE " & SiteCrit & _
                                "ORDER BY RVLID ASC"

                End If

                Overclass.ResetCollection()
                Overclass.CreateDataSet(SqlCode, BindingSource1, ctl)

                ctl.columns("QueryID").visible = False
                ctl.columns("TypeCode").visible = False
                ctl.columns("SiteCode").visible = False
                ctl.columns("RespondCode").visible = False
                ctl.columns("VisitName").readonly = True
                ctl.columns("FormName").readonly = True
                ctl.columns("Status").readonly = True
                ctl.columns("Description").readonly = True
                ctl.columns("RVLID").readonly = True
                ctl.AllowUserToAddRows = False
                Dim cmb As New DataGridViewComboBoxColumn()
                cmb.DataSource = Overclass.TempDataTable("SELECT Code & ' - ' & Site AS Display, Code FROM SiteCode a inner join Study b ON a.ListID=b.CodeList " & _
                                             "WHERE StudyCode='" & Me.ComboBox1.SelectedValue.ToString & "' ORDER BY Code ASC")
                cmb.DataPropertyName = Overclass.CurrentDataSet.Tables(0).Columns(1).ToString
                cmb.ValueMember = "Code"
                cmb.DisplayMember = "Display"
                ctl.Columns.Add(cmb)
                Dim cmb2 As New DataGridViewComboBoxColumn()
                cmb2.DataSource = Overclass.TempDataTable("SELECT Code & ' - ' & ErrorType AS Display, Code FROM TypeCode a inner join Study b ON a.ListID=b.CodeList " & _
                                             "WHERE StudyCode='" & Me.ComboBox1.SelectedValue.ToString & "' ORDER BY Code ASC")
                cmb2.DataPropertyName = Overclass.CurrentDataSet.Tables(0).Columns(2).ToString
                cmb2.ValueMember = "Code"
                cmb2.DisplayMember = "Display"
                ctl.Columns.Add(cmb2)
                Dim cmb3 As New DataGridViewComboBoxColumn()
                cmb3.DataSource = Overclass.TempDataTable("SELECT Code & ' - ' & Group AS Display, Code FROM GroupCode a inner join Study b ON a.ListID=b.CodeList " & _
                                             "WHERE StudyCode='" & Me.ComboBox1.SelectedValue.ToString & "' ORDER BY Code ASC")
                cmb3.DataPropertyName = Overclass.CurrentDataSet.Tables(0).Columns(4).ToString
                cmb3.ValueMember = "Code"
                cmb3.DisplayMember = "Display"
                ctl.Columns.Add(cmb3)
                ctl.columns(3).displayindex = 10
                cmb3.HeaderText = "Respond Code"
                cmb2.HeaderText = "Type Code"
                cmb.HeaderText = "Site Code"
                ctl.columns("FormName").HeaderText = "Form Name"
                ctl.columns("VisitName").HeaderText = "Visit Name"
                ctl.Columns("Person").maxinputlength = 3


        End Select


    End Sub

    Private Sub CheckBox1_CheckStateChanged_1(sender As Object, e As EventArgs) Handles CheckBox1.CheckStateChanged
        Call Specifics(Me.DataGridView2)
    End Sub
End Class