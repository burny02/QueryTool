Public Class AddQuery

    Private Sub AddQuery_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.ComboBox2.DataSource = Overclass.TempDataTable("SELECT DisplayName, StudyCode FROM Study ORDER BY StudyCode ASC")
        Me.ComboBox2.ValueMember = "StudyCode"
        Me.ComboBox2.DisplayMember = "DisplayName"

        Dim SQLCode As String = "SELECT * " & _
                                "FROM Queries " & _
                                "WHERE Study='" & Me.ComboBox2.SelectedValue.ToString & "' " & _
                                "AND Status='Open' " & _
                                "ORDER BY RVLID ASC"

        Overclass.CreateDataSet(SQLCode, Me.BindingSource1, Me.DataGridView3)

        Dim cmb As New DataGridViewImageColumn
        cmb.HeaderText = "Close Query"
        cmb.Image = My.Resources.TICK
        cmb.ImageLayout = DataGridViewImageCellLayout.Zoom
        Me.DataGridView3.Columns.Add(cmb)
        cmb.Name = "CloseQuery"

    End Sub

    Private Sub ComboBox2_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox2.SelectionChangeCommitted

        Dim SQLCode As String = "SELECT * " & _
                                "FROM Queries " & _
                                "WHERE Study='" & Me.ComboBox2.SelectedValue.ToString & "' " & _
                                "AND Status='Open' " & _
                                "ORDER BY RVLID ASC"

        Overclass.CreateDataSet(SQLCode, Me.BindingSource1, Me.DataGridView3)

    End Sub

    Private Sub DataGridView3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick

        If e.ColumnIndex = sender.columns("CloseQuery").index Then

            If IsDBNull(Me.DataGridView3.Item(sender.columns("QueryID").index, e.RowIndex).Value) Then Exit Sub

            'CLOSE THE QUERY
            MsgBox("OK")

        End If

    End Sub

    Private Sub AddQuery_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed

        If AccessLevel = 1 Then
            Overclass.CloseCon()
            Application.Exit()
        End If

    End Sub
End Class