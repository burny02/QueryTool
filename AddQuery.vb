Imports Microsoft.Reporting.WinForms

Public Class AddQuery

    Private Sub AddQuery_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        If Overclass.UnloadData = True Then
            e.Cancel = True
        Else
            RespondCommands.Clear()
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
                NewRow.Item("Priority") = Trim(NewQueryGrid.Item("Priority", e.RowIndex).Value)
                NewRow.Item("Study") = Trim(NewQueryGrid.Item("Study", e.RowIndex).Value)
                NewRow.Item("RVLID") = NewQueryGrid.Item("RVLID", e.RowIndex).Value
                NewRow.Item("Initials") = NewQueryGrid.Item("Initials", e.RowIndex).Value
                NewRow.Item("Status") = "Open"

                Overclass.CurrentDataSet.Tables(0).Rows.Add(NewRow)
                NewQueryGrid.CurrentCell = NewQueryGrid.Item("RVLID", NewQueryGrid.NewRowIndex)

            End If
        End If

        If e.ColumnIndex = sender.columns("StatusCmb").index Then

            If Me.NewQueryGrid.Item("Status", e.RowIndex).Value = "Closed" Then Exit Sub

            If MsgBox("Are you sure you want to close this query?", vbYesNo) = vbYes Then
                Me.NewQueryGrid.Item("ClosedDate", e.RowIndex).Value = Format(DateTime.Now, "dd-MMM-yyyy")
                Me.NewQueryGrid.Item("ClosedTime", e.RowIndex).Value = Format(DateTime.Now, "HH:mm")
                Me.NewQueryGrid.Item("ClosedBy", e.RowIndex).Value = Overclass.GetUserName
                Me.NewQueryGrid.Item("ClosedByRole", e.RowIndex).Value = Role
                Me.NewQueryGrid.Item("Status", e.RowIndex).Value = "Closed"
                BindingSource1.EndEdit()
            End If
        End If

        If e.ColumnIndex = sender.columns("RespondClm").index Then

            If NewQueryGrid.Item("Status", e.RowIndex).Value <> "Responded" Then Exit Sub

            Dim QueryID As String
            QueryID = NewQueryGrid.Item("QueryID", e.RowIndex).Value
            Dim Response = InputBox("Please input response to query", "Query Response")
            If Response = "" Then
                Exit Sub
            Else
                Dim SQL As String
                'NewQueryGrid.Item("Status", e.RowIndex).Value = "Open"
                Dim foundRows() As DataRow
                foundRows = Overclass.CurrentDataSet.Tables(0).Select("QueryID='" & QueryID & "'")
                foundRows(0).Item("Status") = "Open"
                foundRows(0).EndEdit()
                SQL = "INSERT INTO Response(QueryID,Response_Text,Response_Person) " &
                "VALUES ('" & QueryID & "', '" & Response & "', '" & Overclass.GetUserName & "')"

                Dim Cmd As OleDb.OleDbCommand
                Cmd = New OleDb.OleDbCommand(SQL)
                Overclass.SetCommandConnection(Cmd)
                RespondCommands.Add(Cmd)


            End If
        End If

        If e.ColumnIndex = sender.columns("ViewClm").index Then
            If IsDBNull(NewQueryGrid.Item("QueryID", e.RowIndex).Value) = False Then
                Dim QueryID As String
                QueryID = NewQueryGrid.Item("QueryID", e.RowIndex).Value
                Dim CSVString As String = Overclass.CreateCSVString(
                "SELECT format(Response_Timestamp,'dd-MMM-yyyy HH:mm') & ' (' & response_Person & ')  -  ' " &
                "& replace(response_text,',',';') FROM Response WHERE QueryID='" & QueryID & "'")
                CSVString = Replace(CSVString, ",", vbNewLine & vbNewLine)
                If CSVString = "" Then CSVString = "No history found"
                MsgBox(" " & CSVString)
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

        On Error Resume Next

        If String.IsNullOrWhiteSpace(NewQueryGrid.Item("QueryID", e.RowIndex).FormattedValue) Then
            NewQueryGrid.Columns("CreatedBy").ReadOnly = True
            NewQueryGrid.Columns("StatusCmb").ReadOnly = True
            Exit Sub
        End If


        If (NewQueryGrid.Item(sender.columns("Status").index, e.RowIndex).Value) = "Closed" Then
            If e.ColumnIndex <> AdQry.NewQueryGrid.Columns("ResolvedBy").Index And
            e.ColumnIndex <> AdQry.NewQueryGrid.Columns("ResolvedDate").Index Then
                NewQueryGrid.Item(e.ColumnIndex, e.RowIndex).ReadOnly = True
                Exit Sub
            End If
        End If

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

    Private Sub NewQueryGrid_DefaultValuesNeeded(sender As Object, e As DataGridViewRowEventArgs) Handles NewQueryGrid.DefaultValuesNeeded
        NewQueryGrid.Item("Status", e.Row.Index).Value = "Open"
    End Sub

    Private Sub NewQueryGrid_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles NewQueryGrid.DataError

    End Sub

    Private Sub NewQueryGrid_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles NewQueryGrid.CellFormatting

        If e.RowIndex <= -1 Then Exit Sub
        On Error Resume Next
        If e.ColumnIndex = Me.NewQueryGrid.Columns("StatusCmb").Index And Me.NewQueryGrid.Item("Status", e.RowIndex).Value = "Closed" Then e.Value = My.Resources.hyphen
        If e.ColumnIndex = Me.NewQueryGrid.Columns("RespondClm").Index And Me.NewQueryGrid.Item("Status", e.RowIndex).Value <> "Responded" Then e.Value = My.Resources.hyphen
    End Sub
End Class