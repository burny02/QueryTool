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

        Me.NewQueryGrid.Columns.Clear()
        Me.NewQueryGrid2.Columns.Clear()

        Select Case e.TabPageIndex

            Case 0
                Form1.Specifics(NewQueryGrid)

            Case 1
                Form1.Specifics(NewQueryGrid2)


        End Select


    End Sub

    Private Sub AddQuery_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1_Selecting(Me.TabControl1, New TabControlCancelEventArgs(TabPage1, 0, False, TabControlAction.Selecting))

    End Sub

    Private Sub NewQueryGrid_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles NewQueryGrid.CellContentClick

        If e.RowIndex < 0 Then Exit Sub
        If e.ColumnIndex = sender.columns("CloseQuery").index Then

            If IsDBNull(Me.NewQueryGrid.Item(sender.columns("QueryID").index, e.RowIndex).Value) Then
                MsgBox("Please save query first")
                Exit Sub
            End If


            'CLOSE THE QUERY
            If (Me.NewQueryGrid.Item(sender.columns("Status").index, e.RowIndex).Value) = "Closed" Then Exit Sub

            If MsgBox("Are you sure you want to close this query?" & vbNewLine &
                    "Please save to commit changes", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then

                Me.NewQueryGrid.Item("Status", e.RowIndex).Value = "Closed"
                Me.NewQueryGrid.Item("ClosedDate", e.RowIndex).Value = Format(DateTime.Now, "dd-MMM-yyyy")
                Me.NewQueryGrid.Item("ClosedTime", e.RowIndex).Value = Format(DateTime.Now, "HH: mm")
                Me.NewQueryGrid.Item("ClosedBy", e.RowIndex).Value = Overclass.GetUserName
                Me.NewQueryGrid.Item("ClosedByRole", e.RowIndex).Value = Role
                sender.CurrentCell = Nothing
                sender.Rows(e.RowIndex).Visible = False



            End If
        End If

    End Sub

    Private Sub CheckBox201_Click(sender As Object, e As EventArgs) Handles CheckBox201.Click

        If Overclass.UnloadData() = True Then
            RemoveHandler CheckBox201.Click, AddressOf CheckBox201_Click
            CheckBox201.Checked = Not CheckBox201.Checked
            AddHandler CheckBox201.Click, AddressOf CheckBox201_Click
        End If

        Dim TempStudy As String = FilterCombo50.SelectedValue
        Call Form1.Specifics(Me.NewQueryGrid2)
        If TempStudy IsNot Nothing Then FilterCombo50.SelectedValue = TempStudy
        FilterCombo40.SelectedValue = ""
        FilterCombo60.SelectedValue = ""
        FilterCombo70.SelectedValue = ""
        FilterCombo80.SelectedValue = ""

    End Sub

End Class