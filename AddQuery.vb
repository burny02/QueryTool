Imports Microsoft.Reporting.WinForms

Public Class AddQuery

    Private Sub AddQuery_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        If Overclass.UnloadData = True Then
            e.Cancel = True
        Else
            If AccessLevel = 1 Then
                Overclass.CloseCon()
                Application.Exit()
            End If
        End If

    End Sub

    Private Sub TabControl1_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabControl1.Selecting

        Dim SQLCode As String = vbNullString

        If Overclass.UnloadData() = True Then
            e.Cancel = True
            Exit Sub
        End If

        Me.NewQueryGrid.Columns.Clear()
        Me.NewQueryGrid2.Columns.Clear()
        Me.NewQueryGrid3.Columns.Clear()

        Select Case e.TabPageIndex

            Case 0
                Me.ComboBox2.DataSource = Overclass.TempDataTable("SELECT DisplayName, StudyCode FROM Study ORDER BY StudyCode ASC")
                Me.ComboBox2.ValueMember = "StudyCode"
                Me.ComboBox2.DisplayMember = "DisplayName"

                SQLCode = "SELECT QueryID, CreatedBy, Status, Study, FieldName, CreateDate, CreateTime, CreatedByRole, ClosedDate, ClosedTime, " & _
                    "ClosedBy, ClosedByRole, RVLID, Initials, VisitName, FormName, PageNo, Description " & _
                    "FROM Queries " & _
                    "WHERE Study='" & Me.ComboBox2.SelectedValue.ToString & "' " & _
                    "AND QueryID Like 'MANUAL-%' " & _
                    "AND CreatedByRole='" & Role & "' " & _
                    "AND Status='Open' " & _
                    "ORDER BY RVLID ASC"

                Overclass.CreateDataSet(SQLCode, Me.BindingSource1, Me.NewQueryGrid)

                Dim cmb As New DataGridViewImageColumn
                cmb.HeaderText = "Close Query"
                cmb.Image = My.Resources.TICK
                cmb.ImageLayout = DataGridViewImageCellLayout.Zoom
                cmb.Name = "CloseQuery"


                Me.NewQueryGrid.Columns.Add(cmb)

                Me.NewQueryGrid.Columns("QueryID").Visible = False
                Me.NewQueryGrid.Columns("Study").Visible = False
                Me.NewQueryGrid.Columns("FieldName").Visible = False
                Me.NewQueryGrid.Columns("CreateDate").Visible = False
                Me.NewQueryGrid.Columns("CreateTime").Visible = False
                Me.NewQueryGrid.Columns("CreatedByRole").Visible = False
                Me.NewQueryGrid.Columns("ClosedDate").Visible = False
                Me.NewQueryGrid.Columns("ClosedTime").Visible = False
                Me.NewQueryGrid.Columns("ClosedBy").Visible = False
                Me.NewQueryGrid.Columns("ClosedByRole").Visible = False

                Me.NewQueryGrid.Columns("Status").ReadOnly = True
                Me.NewQueryGrid.Columns("CreatedBy").ReadOnly = True

                Me.NewQueryGrid.Columns("VisitName").HeaderText = "Study Visit"
                Me.NewQueryGrid.Columns("FormName").HeaderText = "Assessment/Procedure"
                Me.NewQueryGrid.Columns("PageNo").HeaderText = "Page No."
                Me.NewQueryGrid.Columns("Description").HeaderText = "Query Description"
                Me.NewQueryGrid.Columns("CreatedBy").HeaderText = "Created By"

            Case 1

                Me.ComboBox1.DataSource = Overclass.TempDataTable("SELECT DisplayName, StudyCode FROM Study ORDER BY StudyCode ASC")
                Me.ComboBox1.ValueMember = "StudyCode"
                Me.ComboBox1.DisplayMember = "DisplayName"

                Dim AllowedSite As String = Overclass.CreateCSVString("SELECT Code FROM SiteCODE a INNER JOIN Study b ON a.ListID=b.CodeList " & _
                                                            "WHERE StudyCode='" & Me.ComboBox1.SelectedValue.ToString & "'")
                Dim AllowedResponse As String = Overclass.CreateCSVString("SELECT Code FROM GroupCode a INNER JOIN Study b ON a.ListID=b.CodeList " & _
                                                            "WHERE StudyCode='" & Me.ComboBox1.SelectedValue.ToString & "'")
                Dim AllowedType As String = Overclass.CreateCSVString("SELECT Code FROM TypeCode a INNER JOIN Study b ON a.ListID=b.CodeList " & _
                                                            "WHERE StudyCode='" & Me.ComboBox1.SelectedValue.ToString & "'")

                SQLCode = "SELECT a.QueryID, CreatedBy, RVLID, VisitName, FormName, SiteCode, TypeCode, Person, RespondCode, " & _
                                "Description FROM QueryCodes as a INNER JOIN Queries as b ON a.QueryID=b.QueryID " & _
                                "WHERE Study='" & Me.ComboBox1.SelectedValue.ToString & "'" & _
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

                Overclass.CreateDataSet(SQLCode, Me.BindingSource1, Me.NewQueryGrid2)
                Me.NewQueryGrid2.Columns("QueryID").Visible = False
                Me.NewQueryGrid2.Columns("SiteCode").Visible = False
                Me.NewQueryGrid2.Columns("TypeCode").Visible = False
                Me.NewQueryGrid2.Columns("RespondCode").Visible = False
                Me.NewQueryGrid2.Columns("RVLID").ReadOnly = True
                Me.NewQueryGrid2.Columns("FormName").ReadOnly = True
                Me.NewQueryGrid2.Columns("Description").ReadOnly = True
                Me.NewQueryGrid2.Columns("VisitName").ReadOnly = True
                Me.NewQueryGrid2.Columns("FormName").ReadOnly = True
                Me.NewQueryGrid2.Columns("CreatedBy").ReadOnly = True

                Me.NewQueryGrid2.AllowUserToAddRows = False
                Dim cmb As New DataGridViewComboBoxColumn()
                cmb.DataSource = Overclass.TempDataTable("SELECT Code & ' - ' & Site AS Display, Code FROM SiteCode a inner join Study b ON a.ListID=b.CodeList " & _
                                             "WHERE StudyCode='" & Me.ComboBox1.SelectedValue.ToString & "' ORDER BY Code ASC")
                cmb.DataPropertyName = Overclass.CurrentDataSet.Tables(0).Columns("SiteCode").ToString
                cmb.ValueMember = "Code"
                cmb.DisplayMember = "Display"
                Me.NewQueryGrid2.Columns.Add(cmb)
                Dim cmb2 As New DataGridViewComboBoxColumn()
                cmb2.DataSource = Overclass.TempDataTable("SELECT Code & ' - ' & ErrorType AS Display, Code FROM TypeCode a inner join Study b ON a.ListID=b.CodeList " & _
                                             "WHERE StudyCode='" & Me.ComboBox1.SelectedValue.ToString & "' ORDER BY Code ASC")
                cmb2.DataPropertyName = Overclass.CurrentDataSet.Tables(0).Columns("TypeCode").ToString
                cmb2.ValueMember = "Code"
                cmb2.DisplayMember = "Display"
                Me.NewQueryGrid2.Columns.Add(cmb2)
                Dim cmb3 As New DataGridViewComboBoxColumn()
                cmb3.DataSource = Overclass.TempDataTable("SELECT Code & ' - ' & Group AS Display, Code FROM GroupCode a inner join Study b ON a.ListID=b.CodeList " & _
                                             "WHERE StudyCode='" & Me.ComboBox1.SelectedValue.ToString & "' ORDER BY Code ASC")
                cmb3.DataPropertyName = Overclass.CurrentDataSet.Tables(0).Columns("RespondCode").ToString
                cmb3.ValueMember = "Code"
                cmb3.DisplayMember = "Display"
                Me.NewQueryGrid2.Columns.Add(cmb3)
                cmb3.HeaderText = "Respond Code"
                cmb2.HeaderText = "Type Code"
                cmb.HeaderText = "Site Code"
                Me.NewQueryGrid2.Columns("FormName").HeaderText = "Form Name"
                Dim ctl As Object = Me.NewQueryGrid2
                ctl.Columns("Person").maxinputlength = 3
                Me.NewQueryGrid2.Columns("Person").DisplayIndex = Me.NewQueryGrid2.Columns.Count - 2

            Case 2

                Me.ComboBox3.DataSource = Overclass.TempDataTable("SELECT DisplayName, StudyCode FROM Study ORDER BY StudyCode ASC")
                Me.ComboBox3.ValueMember = "StudyCode"
                Me.ComboBox3.DisplayMember = "DisplayName"

                SQLCode = "SELECT a.QueryID, CreatedBy, RVLID, VisitName, FormName, SiteCode, TypeCode, Person, RespondCode, " & _
                                "Description FROM QueryCodes as a INNER JOIN Queries as b ON a.QueryID=b.QueryID " & _
                                "WHERE Study='" & Me.ComboBox3.SelectedValue.ToString & "' AND a.QueryID Like 'MANUAL-%' " & _
                                "AND CreatedByRole='" & Role & "'" & _
                                "ORDER BY RVLID ASC"


                Overclass.CreateDataSet(SQLCode, Me.BindingSource1, Me.NewQueryGrid3)

                Me.NewQueryGrid3.Columns("QueryID").Visible = False
                Me.NewQueryGrid3.Columns("SiteCode").Visible = False
                Me.NewQueryGrid3.Columns("TypeCode").Visible = False
                Me.NewQueryGrid3.Columns("RespondCode").Visible = False
                Me.NewQueryGrid3.Columns("RVLID").ReadOnly = True
                Me.NewQueryGrid3.Columns("FormName").ReadOnly = True
                Me.NewQueryGrid3.Columns("Description").ReadOnly = True
                Me.NewQueryGrid3.Columns("VisitName").ReadOnly = True
                Me.NewQueryGrid3.Columns("FormName").ReadOnly = True
                Me.NewQueryGrid3.Columns("CreatedBy").ReadOnly = True

                Dim cmb As New DataGridViewComboBoxColumn()
                cmb.DataSource = Overclass.TempDataTable("SELECT Code & ' - ' & Site AS Display, Code FROM SiteCode a inner join Study b ON a.ListID=b.CodeList " & _
                                             "WHERE StudyCode='" & Me.ComboBox3.SelectedValue.ToString & "' ORDER BY Code ASC")
                cmb.DataPropertyName = Overclass.CurrentDataSet.Tables(0).Columns("SiteCode").ToString
                cmb.ValueMember = "Code"
                cmb.DisplayMember = "Display"
                Me.NewQueryGrid3.Columns.Add(cmb)
                Dim cmb2 As New DataGridViewComboBoxColumn()
                cmb2.DataSource = Overclass.TempDataTable("SELECT Code & ' - ' & ErrorType AS Display, Code FROM TypeCode a inner join Study b ON a.ListID=b.CodeList " & _
                                             "WHERE StudyCode='" & Me.ComboBox3.SelectedValue.ToString & "' ORDER BY Code ASC")
                cmb2.DataPropertyName = Overclass.CurrentDataSet.Tables(0).Columns("TypeCode").ToString
                cmb2.ValueMember = "Code"
                cmb2.DisplayMember = "Display"
                Me.NewQueryGrid3.Columns.Add(cmb2)
                Dim cmb3 As New DataGridViewComboBoxColumn()
                cmb3.DataSource = Overclass.TempDataTable("SELECT Code & ' - ' & Group AS Display, Code FROM GroupCode a inner join Study b ON a.ListID=b.CodeList " & _
                                             "WHERE StudyCode='" & Me.ComboBox3.SelectedValue.ToString & "' ORDER BY Code ASC")
                cmb3.DataPropertyName = Overclass.CurrentDataSet.Tables(0).Columns("RespondCode").ToString
                cmb3.ValueMember = "Code"
                cmb3.DisplayMember = "Display"
                Me.NewQueryGrid3.Columns.Add(cmb3)
                cmb3.HeaderText = "Respond Code"
                cmb2.HeaderText = "Type Code"
                cmb.HeaderText = "Site Code"
                Me.NewQueryGrid3.Columns("FormName").HeaderText = "Form Name"
                Dim ctl As Object = Me.NewQueryGrid3
                ctl.Columns("Person").maxinputlength = 3
                Me.NewQueryGrid3.Columns("Person").DisplayIndex = Me.NewQueryGrid3.Columns.Count - 2

        End Select


    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        Call Saver(Me.NewQueryGrid)
    End Sub

    Private Sub AddQuery_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1_Selecting(Me.TabControl1, New TabControlCancelEventArgs(TabPage1, 0, False, TabControlAction.Selecting))

    End Sub

    Private Sub NewQueryGrid_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles NewQueryGrid.CellContentClick

        If e.ColumnIndex = sender.columns("CloseQuery").index Then

            If IsDBNull(Me.NewQueryGrid.Item(sender.columns("QueryID").index, e.RowIndex).Value) Then
                MsgBox("Please save query first")
                Exit Sub
            End If


            'CLOSE THE QUERY
            If (Me.NewQueryGrid.Item(sender.columns("Status").index, e.RowIndex).Value) = "Closed" Then Exit Sub

            If MsgBox("Are you sure you want to close this query?" & vbNewLine & _
                    "Please save to commit changes", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then

                Me.NewQueryGrid.Item("Status", e.RowIndex).Value = "Closed"
                Me.NewQueryGrid.Item("ClosedDate", e.RowIndex).Value = Format(DateTime.Now, "dd-MMM-yyyy")
                Me.NewQueryGrid.Item("ClosedTime", e.RowIndex).Value = Format(DateTime.Now, "HH:mm")
                Me.NewQueryGrid.Item("ClosedBy", e.RowIndex).Value = Overclass.GetUserName
                Me.NewQueryGrid.Item("ClosedByRole", e.RowIndex).Value = Role
                sender.CurrentCell = Nothing
                sender.Rows(e.RowIndex).Visible = False



            End If
        End If

    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox1.SelectionChangeCommitted

        If Overclass.UnloadData = True Then
            Exit Sub
        End If


        Dim AllowedSite As String = Overclass.CreateCSVString("SELECT Code FROM SiteCODE a INNER JOIN Study b ON a.ListID=b.CodeList " & _
                                                            "WHERE StudyCode='" & Me.ComboBox1.SelectedValue.ToString & "'")
        Dim AllowedResponse As String = Overclass.CreateCSVString("SELECT Code FROM GroupCode a INNER JOIN Study b ON a.ListID=b.CodeList " & _
                                                    "WHERE StudyCode='" & Me.ComboBox1.SelectedValue.ToString & "'")
        Dim AllowedType As String = Overclass.CreateCSVString("SELECT Code FROM TypeCode a INNER JOIN Study b ON a.ListID=b.CodeList " & _
                                                    "WHERE StudyCode='" & Me.ComboBox1.SelectedValue.ToString & "'")

        Dim SQLCode As String = "SELECT a.QueryID, CreatedBy, RVLID, VisitName, FormName, SiteCode, TypeCode, Person, RespondCode, " & _
                                "Description FROM QueryCodes as a INNER JOIN Queries as b ON a.QueryID=b.QueryID " & _
                                "WHERE Study='" & Me.ComboBox1.SelectedValue.ToString & "'" & _
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

        Overclass.CreateDataSet(SQLCode, Me.BindingSource1, Me.NewQueryGrid2)


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call Saver(Me.NewQueryGrid2)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call Saver(Me.NewQueryGrid3)
    End Sub

    Private Sub ComboBox3_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox3.SelectionChangeCommitted

        If Overclass.UnloadData = True Then
            Exit Sub
        End If

        Dim SQLCode As String = "SELECT a.QueryID, CreatedBy, RVLID, VisitName, FormName, SiteCode, TypeCode, Person, RespondCode, " & _
                                "Description FROM QueryCodes as a INNER JOIN Queries as b ON a.QueryID=b.QueryID " & _
                                "WHERE Study='" & Me.ComboBox3.SelectedValue.ToString & "' AND a.QueryID Like 'MANUAL-%' " & _
                                "AND CreatedByRole='" & Role & "'" & _
                                "ORDER BY RVLID ASC"


        Overclass.CreateDataSet(SQLCode, Me.BindingSource1, Me.NewQueryGrid3)

    End Sub

    Private Sub ComboBox2_SelectionChangeCommitted_1(sender As Object, e As EventArgs) Handles ComboBox2.SelectionChangeCommitted

        If Overclass.UnloadData = True Then
            Exit Sub
        End If

        Dim SQLCode As String = "SELECT QueryID, CreatedBy, Status, Study, FieldName, CreateDate, CreateTime, CreatedBy, CreatedByRole, ClosedDate, ClosedTime, " & _
            "ClosedBy, ClosedByRole, RVLID, Initials, VisitName, FormName, PageNo, Description " & _
            "FROM Queries " & _
            "WHERE Study='" & Me.ComboBox2.SelectedValue.ToString & "' " & _
            "AND QueryID Like 'MANUAL-%' " & _
            "AND CreatedByRole='" & Role & "' " & _
            "AND Status='Open' " & _
            "ORDER BY RVLID ASC"

        Overclass.CreateDataSet(SQLCode, Me.BindingSource1, Me.NewQueryGrid)

        Me.NewQueryGrid.Columns("CloseQuery").DisplayIndex = Me.NewQueryGrid.Columns.Count - 1

    End Sub

    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click

        MsgBox("Only correctly allocated queries will print out. Please ensure queries are saved first")


        Dim RVLID As Long = 0
        Dim InputString As String = vbNullString

        InputString = InputBox("Please input RVLID to print", "RVLID", "123456")

        If InputString = vbNullString Then Exit Sub

        Try
            RVLID = CLng(InputString)
        Catch ex As Exception
            Exit Sub
        End Try

        'SELECT RVLID FROM CURRENTDATA SET

        Try

            Dim SqlString As String = "SELECT * FROM PrintOut WHERE RVLID='" & RVLID & "'" & _
                                                          " AND CreatedByRole='" & Role & "'" & _
                                                          " AND Status='Open'"

            Dim dt As DataTable = Overclass.TempDataTable(SqlString)

            If dt.Rows.Count = 0 Then
                MsgBox("No queries found for volunteer " & RVLID)
                Exit Sub
            End If

            'RUN REPORT SEPERATING BY VISIT 

            Dim OK As New ReportViewer
            OK.Visible = True
            OK.ReportViewer1.Visible = True
            OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "QueryTool.PrintReport.rdlc"
            OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet", _
                                                       dt))
            OK.ReportViewer1.RefreshReport()

        Catch ex As Exception

            MsgBox(ex.Message)
            Exit Sub

        End Try





    End Sub
End Class