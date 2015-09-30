Public Class AddQuery

    Private Sub ComboBox2_SelectionChangeCommitted(sender As Object, e As EventArgs)

        If Overclass.UnloadData = True Then
            Exit Sub
        End If

        Dim SQLCode As String = "SELECT QueryID, Study, FieldName, CreateDate, CreateTime, CreatedBy, CreatedByRole, ClosedDate, ClosedTime, " & _
            "ClosedBy, ClosedByRole, RVLID, Initials, VisitName, FormName, Status, PageNo, Description " & _
            "FROM Queries " & _
            "WHERE Study='" & Me.ComboBox2.SelectedValue.ToString & "' " & _
            "AND QueryID Like 'MANUAL-%' " & _
            "ORDER BY RVLID ASC"

        Overclass.CreateDataSet(SQLCode, Me.BindingSource1, Me.NewQueryGrid)


    End Sub

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

        Select Case e.TabPageIndex

            Case 0
                Me.ComboBox2.DataSource = Overclass.TempDataTable("SELECT DisplayName, StudyCode FROM Study ORDER BY StudyCode ASC")
                Me.ComboBox2.ValueMember = "StudyCode"
                Me.ComboBox2.DisplayMember = "DisplayName"

                SQLCode = "SELECT QueryID, Study, FieldName, CreateDate, CreateTime, CreatedBy, CreatedByRole, ClosedDate, ClosedTime, " & _
                    "ClosedBy, ClosedByRole, RVLID, Initials, VisitName, FormName, Status, PageNo, Description " & _
                    "FROM Queries " & _
                    "WHERE Study='" & Me.ComboBox2.SelectedValue.ToString & "' " & _
                    "AND QueryID Like 'MANUAL-%' " & _
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
                Me.NewQueryGrid.Columns("CreatedBy").Visible = False
                Me.NewQueryGrid.Columns("CreatedByRole").Visible = False
                Me.NewQueryGrid.Columns("ClosedDate").Visible = False
                Me.NewQueryGrid.Columns("ClosedTime").Visible = False
                Me.NewQueryGrid.Columns("ClosedBy").Visible = False
                Me.NewQueryGrid.Columns("ClosedByRole").Visible = False

                Me.NewQueryGrid.Columns("Status").ReadOnly = True

                Me.NewQueryGrid.Columns("VisitName").HeaderText = "Study Visit"
                Me.NewQueryGrid.Columns("FormName").HeaderText = "Assessment/Procedure"
                Me.NewQueryGrid.Columns("PageNo").HeaderText = "Page No."
                Me.NewQueryGrid.Columns("Description").HeaderText = "Query Description"

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

                SQLCode = "SELECT a.QueryID, SiteCode, TypeCode, Person, RespondCode, RVLID, " & _
                                "FormName, Description, Status FROM QueryCodes as a INNER JOIN Queries as b ON a.QueryID=b.QueryID " & _
                                "WHERE Study='" & Me.ComboBox1.SelectedValue.ToString & "'" & _
                                " AND a.QueryID Like 'MANUAL-%' " & _
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
                Me.NewQueryGrid2.Columns(0).Visible = False
                Me.NewQueryGrid2.Columns(1).Visible = False
                Me.NewQueryGrid2.Columns(2).Visible = False
                Me.NewQueryGrid2.Columns(4).Visible = False
                Me.NewQueryGrid2.Columns(5).ReadOnly = True
                Me.NewQueryGrid2.Columns(6).ReadOnly = True
                Me.NewQueryGrid2.Columns(7).ReadOnly = True
                Me.NewQueryGrid2.Columns(8).ReadOnly = True
                Me.NewQueryGrid2.AllowUserToAddRows = False
                Dim cmb As New DataGridViewComboBoxColumn()
                cmb.DataSource = Overclass.TempDataTable("SELECT Code & ' - ' & Site AS Display, Code FROM SiteCode a inner join Study b ON a.ListID=b.CodeList " & _
                                             "WHERE StudyCode='" & Me.ComboBox1.SelectedValue.ToString & "' ORDER BY Code ASC")
                cmb.DataPropertyName = Overclass.CurrentDataSet.Tables(0).Columns(1).ToString
                cmb.ValueMember = "Code"
                cmb.DisplayMember = "Display"
                Me.NewQueryGrid2.Columns.Add(cmb)
                Dim cmb2 As New DataGridViewComboBoxColumn()
                cmb2.DataSource = Overclass.TempDataTable("SELECT Code & ' - ' & ErrorType AS Display, Code FROM TypeCode a inner join Study b ON a.ListID=b.CodeList " & _
                                             "WHERE StudyCode='" & Me.ComboBox1.SelectedValue.ToString & "' ORDER BY Code ASC")
                cmb2.DataPropertyName = Overclass.CurrentDataSet.Tables(0).Columns(2).ToString
                cmb2.ValueMember = "Code"
                cmb2.DisplayMember = "Display"
                Me.NewQueryGrid2.Columns.Add(cmb2)
                Dim cmb3 As New DataGridViewComboBoxColumn()
                cmb3.DataSource = Overclass.TempDataTable("SELECT Code & ' - ' & Group AS Display, Code FROM GroupCode a inner join Study b ON a.ListID=b.CodeList " & _
                                             "WHERE StudyCode='" & Me.ComboBox1.SelectedValue.ToString & "' ORDER BY Code ASC")
                cmb3.DataPropertyName = Overclass.CurrentDataSet.Tables(0).Columns(4).ToString
                cmb3.ValueMember = "Code"
                cmb3.DisplayMember = "Display"
                Me.NewQueryGrid2.Columns.Add(cmb3)
                Me.NewQueryGrid2.Columns(3).DisplayIndex = 10
                cmb3.HeaderText = "Respond Code"
                cmb2.HeaderText = "Type Code"
                cmb.HeaderText = "Site Code"
                Me.NewQueryGrid2.Columns(6).HeaderText = "Form Name"
                Dim ctl As Object = Me.NewQueryGrid2
                ctl.Columns(3).maxinputlength = 3

            Case 2

                Me.ComboBox3.DataSource = Overclass.TempDataTable("SELECT DisplayName, StudyCode FROM Study ORDER BY StudyCode ASC")
                Me.ComboBox3.ValueMember = "StudyCode"
                Me.ComboBox3.DisplayMember = "DisplayName"

                SQLCode = "SELECT a.QueryID, SiteCode, TypeCode, Person, RespondCode, RVLID, " & _
                                "FormName, Description, Status FROM QueryCodes as a INNER JOIN Queries as b ON a.QueryID=b.QueryID " & _
                                "WHERE Study='" & Me.ComboBox3.SelectedValue.ToString & "' AND a.QueryID Like 'MANUAL-%' " & _
                                "ORDER BY RVLID ASC"


                Overclass.CreateDataSet(SQLCode, Me.BindingSource1, Me.NewQueryGrid3)

                Me.NewQueryGrid3.Columns(0).Visible = False
                Me.NewQueryGrid3.Columns(1).Visible = False
                Me.NewQueryGrid3.Columns(2).Visible = False
                Me.NewQueryGrid3.Columns(4).Visible = False
                Me.NewQueryGrid3.Columns(5).ReadOnly = True
                Me.NewQueryGrid3.Columns(6).ReadOnly = True
                Me.NewQueryGrid3.Columns(7).ReadOnly = True
                Me.NewQueryGrid3.Columns(8).ReadOnly = True
                Me.NewQueryGrid3.AllowUserToAddRows = False
                Dim cmb As New DataGridViewComboBoxColumn()
                cmb.DataSource = Overclass.TempDataTable("SELECT Code & ' - ' & Site AS Display, Code FROM SiteCode a inner join Study b ON a.ListID=b.CodeList " & _
                                             "WHERE StudyCode='" & Me.ComboBox1.SelectedValue.ToString & "' ORDER BY Code ASC")
                cmb.DataPropertyName = Overclass.CurrentDataSet.Tables(0).Columns(1).ToString
                cmb.ValueMember = "Code"
                cmb.DisplayMember = "Display"
                Me.NewQueryGrid3.Columns.Add(cmb)
                Dim cmb2 As New DataGridViewComboBoxColumn()
                cmb2.DataSource = Overclass.TempDataTable("SELECT Code & ' - ' & ErrorType AS Display, Code FROM TypeCode a inner join Study b ON a.ListID=b.CodeList " & _
                                             "WHERE StudyCode='" & Me.ComboBox1.SelectedValue.ToString & "' ORDER BY Code ASC")
                cmb2.DataPropertyName = Overclass.CurrentDataSet.Tables(0).Columns(2).ToString
                cmb2.ValueMember = "Code"
                cmb2.DisplayMember = "Display"
                Me.NewQueryGrid3.Columns.Add(cmb2)
                Dim cmb3 As New DataGridViewComboBoxColumn()
                cmb3.DataSource = Overclass.TempDataTable("SELECT Code & ' - ' & Group AS Display, Code FROM GroupCode a inner join Study b ON a.ListID=b.CodeList " & _
                                             "WHERE StudyCode='" & Me.ComboBox1.SelectedValue.ToString & "' ORDER BY Code ASC")
                cmb3.DataPropertyName = Overclass.CurrentDataSet.Tables(0).Columns(4).ToString
                cmb3.ValueMember = "Code"
                cmb3.DisplayMember = "Display"
                Me.NewQueryGrid3.Columns.Add(cmb3)
                Me.NewQueryGrid3.Columns(3).DisplayIndex = 10
                cmb3.HeaderText = "Respond Code"
                cmb2.HeaderText = "Type Code"
                cmb.HeaderText = "Site Code"
                Me.NewQueryGrid3.Columns(6).HeaderText = "Form Name"
                Dim ctl As Object = Me.NewQueryGrid3
                ctl.Columns(3).maxinputlength = 3

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

        Dim SQLCode As String = "SELECT a.QueryID, SiteCode, TypeCode, Person, RespondCode, RVLID, " & _
                        "FormName, Description, Status FROM QueryCodes as a INNER JOIN Queries as b ON a.QueryID=b.QueryID " & _
                        "WHERE Study='" & Me.ComboBox1.SelectedValue.ToString & "'" & _
                        " AND a.QueryID Like 'MANUAL-%' " & _
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

        Dim SQLCode As String = "SELECT a.QueryID, SiteCode, TypeCode, Person, RespondCode, RVLID, " & _
                                "FormName, Description, Status FROM QueryCodes as a INNER JOIN Queries as b ON a.QueryID=b.QueryID " & _
                                "WHERE Study='" & Me.ComboBox3.SelectedValue.ToString & "' AND a.QueryID Like 'MANUAL-%' " & _
                                "ORDER BY RVLID ASC"


        Overclass.CreateDataSet(SQLCode, Me.BindingSource1, Me.NewQueryGrid3)

    End Sub
End Class