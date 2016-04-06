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
        If e.ColumnIndex = sender.columns("CopyQuery").index Then
            If MsgBox("Do you want to copy this query to a new line?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then

                Dim NewRow As DataRow = Overclass.CurrentDataSet.Tables(0).NewRow
                NewRow.Item("FieldName") = NewQueryGrid.Item("FieldName", e.RowIndex).Value
                NewRow.Item("VisitName") = NewQueryGrid.Item("VisitName", e.RowIndex).Value
                NewRow.Item("FormName") = NewQueryGrid.Item("FormName", e.RowIndex).Value
                NewRow.Item("PageNo") = NewQueryGrid.Item("PageNo", e.RowIndex).Value
                NewRow.Item("Description") = NewQueryGrid.Item("Description", e.RowIndex).Value
                NewRow.Item("Priority") = NewQueryGrid.Item("Priority", e.RowIndex).Value
                NewRow.Item("Study") = NewQueryGrid.Item("Study", e.RowIndex).Value
                NewRow.Item("RVLID") = NewQueryGrid.Item("RVLID", e.RowIndex).Value
                NewRow.Item("Initials") = NewQueryGrid.Item("Initials", e.RowIndex).Value
                NewRow.Item("Status") = "Open"

                Overclass.CurrentDataSet.Tables(0).Rows.Add(NewRow)
                NewQueryGrid.CurrentCell = NewQueryGrid.Item("RVLID", NewQueryGrid.NewRowIndex)

            End If
        End If
        If e.ColumnIndex = sender.columns("CloseQuery").index Then
            If e.RowIndex = NewQueryGrid.NewRowIndex Then Exit Sub

            If IsDBNull(Me.NewQueryGrid.Item(sender.columns("QueryID").index, e.RowIndex).Value) Then
                If MsgBox("Do you want to delete this query? " & vbNewLine & vbNewLine & "To close the query please save it first", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    NewQueryGrid.Rows.RemoveAt(e.RowIndex)
                    Try
                        NewQueryGrid.CurrentCell = NewQueryGrid.Item("RVLID", e.RowIndex - 1)
                    Catch ex As Exception
                        Try
                            NewQueryGrid.CurrentCell = NewQueryGrid.Item("RVLID", e.RowIndex + 1)
                        Catch ex2 As Exception
                            NewQueryGrid.CurrentCell = NewQueryGrid.Item("RVLID", NewQueryGrid.NewRowIndex)
                        End Try
                    End Try

                End If
                Exit Sub
                End If


                'CLOSE THE QUERY
                If (Me.NewQueryGrid.Item(sender.columns("Status").index, e.RowIndex).Value) = "Closed" Then Exit Sub

                If MsgBox("Are you sure you want to close this query?" & vbNewLine &
                        "Please save to commit changes", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then

                Me.NewQueryGrid.Item("Status", e.RowIndex).Value = "Closed"
                Me.NewQueryGrid.Item("ClosedDate", e.RowIndex).Value = Format(DateTime.Now, "dd-MMM-yyyy")
                Me.NewQueryGrid.Item("ClosedTime", e.RowIndex).Value = Format(DateTime.Now, "HH:mm")
                Me.NewQueryGrid.Item("ClosedBy", e.RowIndex).Value = Overclass.GetUserName
                Me.NewQueryGrid.Item("ClosedByRole", e.RowIndex).Value = Role
                sender.CurrentCell = Nothing
                sender.Rows(e.RowIndex).Visible = False
                Try
                    NewQueryGrid.CurrentCell = NewQueryGrid.Item("RVLID", e.RowIndex - 1)
                Catch ex As Exception
                    Try
                        NewQueryGrid.CurrentCell = NewQueryGrid.Item("RVLID", e.RowIndex + 1)
                    Catch ex2 As Exception
                        NewQueryGrid.CurrentCell = NewQueryGrid.Item("RVLID", NewQueryGrid.NewRowIndex)
                    End Try
                End Try

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

    Private Sub NewQueryGrid_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles NewQueryGrid.CellEnter

        If String.IsNullOrWhiteSpace(AdQry.NewQueryGrid.Item("QueryID", e.RowIndex).FormattedValue) Then
            NewQueryGrid.Columns("CreatedBy").ReadOnly = True
            Exit Sub
        End If

        On Error Resume Next
        If e.ColumnIndex = AdQry.NewQueryGrid.Columns("RVLID").Index Or
            e.ColumnIndex = AdQry.NewQueryGrid.Columns("VisitName").Index Or
            e.ColumnIndex = AdQry.NewQueryGrid.Columns("FormName").Index Or
            e.ColumnIndex = AdQry.NewQueryGrid.Columns("PageNo").Index Or
            e.ColumnIndex = AdQry.NewQueryGrid.Columns("Description").Index Or
            e.ColumnIndex = AdQry.NewQueryGrid.Columns("Initials").Index Or
            e.ColumnIndex = AdQry.NewQueryGrid.Columns("PriorityClm").Index Or
            e.ColumnIndex = AdQry.NewQueryGrid.Columns("CreatedBy").Index Then

            AdQry.NewQueryGrid.Item(e.ColumnIndex, e.RowIndex).ReadOnly = True
        Else
            AdQry.NewQueryGrid.Item(e.ColumnIndex, e.RowIndex).ReadOnly = False
        End If


    End Sub

    Private Sub NewQueryGrid_RowPostPaint(sender As Object, e As DataGridViewRowPostPaintEventArgs) Handles NewQueryGrid.RowPostPaint

        If NewQueryGrid.Item("Status", e.RowIndex).Value = "Closed" Then
            NewQueryGrid.Rows(e.RowIndex).Visible = False
        End If

    End Sub
End Class