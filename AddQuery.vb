Public Class AddQuery

    Private Sub AddQuery_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.ComboBox2.DataSource = Overclass.TempDataTable("SELECT DisplayName, StudyCode FROM Study ORDER BY StudyCode ASC")
        Me.ComboBox2.ValueMember = "StudyCode"
        Me.ComboBox2.DisplayMember = "DisplayName"

        Dim SQLCode As String = "SELECT QueryID, Study, FieldName, CreateDate, CreateTime, CreatedBy, CreatedByRole, ClosedDate, ClosedTime, " & _
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
        Dim cmb2 As New DataGridViewImageColumn
        cmb2.HeaderText = "Assign Query"
        cmb2.Image = My.Resources.TICK
        cmb2.ImageLayout = DataGridViewImageCellLayout.Zoom
        cmb2.Name = "AssignQry"

        Me.NewQueryGrid.Columns.Add(cmb2)
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

    End Sub

    Private Sub ComboBox2_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox2.SelectionChangeCommitted

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

    Private Sub DataGridView3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles NewQueryGrid.CellContentClick

        If e.ColumnIndex = sender.columns("CloseQuery").index Then

            If IsDBNull(Me.NewQueryGrid.Item(sender.columns("QueryID").index, e.RowIndex).Value) Then
                MsgBox("Please save query first")
                Exit Sub
            End If


            'CLOSE THE QUERY
            If MsgBox("Are you sure you want to close this query?" & vbNewLine & _
                    "Please save to commit changes", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then

                Me.NewQueryGrid.Item("Status", e.RowIndex).Value = "Closed"
                Me.NewQueryGrid.Item("ClosedDate", e.RowIndex).Value = Format(DateTime.Now, "dd-MMM-yyyy")
                Me.NewQueryGrid.Item("ClosedTime", e.RowIndex).Value = Format(DateTime.Now, "HH:mm")
                Me.NewQueryGrid.Item("ClosedBy", e.RowIndex).Value = Overclass.GetUserName
                Me.NewQueryGrid.Item("ClosedByRole", e.RowIndex).Value = Role

            End If
        End If

        If e.ColumnIndex = sender.columns("AssignQry").index Then

            If IsDBNull(Me.NewQueryGrid.Item(sender.columns("QueryID").index, e.RowIndex).Value) Then
                MsgBox("Please save query first")
                Exit Sub
            End If


            'Open Assignpage
            MsgBox("OK2")

        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call Saver(Me.NewQueryGrid)
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
End Class