Public Class ResponseView
    Private Timeout As Integer = 0

    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles StaffQueryGrid.CellContentDoubleClick

        If e.RowIndex < 0 Then Exit Sub

        Timeout = 0
        Dim QueryID As String


        If e.ColumnIndex = StaffQueryGrid.Columns("RespondClm").Index Then
            QueryID = StaffQueryGrid.Item("QueryID", e.RowIndex).Value

            Dim RespondText As String = vbNullString

            Dim CSVString As String = Overclass.CreateCSVString(
            "SELECT format(Response_Timestamp,'dd-MMM-yyyy HH:mm') & ' (' & response_Person & ')  -  ' " &
            "& replace(response_text,',',';') FROM Response WHERE QueryID=" & QueryID)
            CSVString = Replace(CSVString, ",", vbNewLine & vbNewLine)
            If CSVString = "" Then CSVString = "No history found"
            RespondText = CSVString

            RespondText = RespondText & vbNewLine & "Please input response to query:"

            Dim Response = InputBox(RespondText, "Query Response", "Data corrected")
            If Response = "" Then
                Exit Sub
            Else
                Response = Replace(Response, "'", "")
                Dim SQL(1) As String
                SQL(0) = "UPDATE Queries SET Status='Responded' WHERE QueryID=" & QueryID
                SQL(1) = "INSERT INTO Response(QueryID,Response_Text,Response_Person) VALUES (" & QueryID & ", '" & Response & "', '" & Overclass.GetUserName & "')"
                Overclass.AddToMassSQL(SQL(0))
                Overclass.AddToMassSQL(SQL(1))
                Overclass.ExecuteMassSQL()
                Overclass.Refresher(StaffQueryGrid)
            End If
        End If

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick


        Timeout += 1
        If Timeout = 30 Then
            Me.Close()
            Exit Sub
        End If
        CurrentRecord = StaffQueryGrid.FirstDisplayedScrollingRowIndex
        Form1.Specifics(Me)
        Try
            StaffQueryGrid.FirstDisplayedScrollingRowIndex = CurrentRecord
        Catch ex As Exception
        End Try
    End Sub

    Private Sub StaffQueryGrid_RowPrePaint(sender As Object, e As DataGridViewRowPrePaintEventArgs) Handles StaffQueryGrid.RowPrePaint

        Try
            Dim RValue As DataGridViewCell = StaffQueryGrid.Rows(e.RowIndex).Cells("RespondClm")
            Dim LValue As DataGridViewCell = StaffQueryGrid.Rows(e.RowIndex).Cells("LightClm")

            If RValue.Tag <> "Done" Then
                If StaffQueryGrid.Item("Bounced", e.RowIndex).Value = True Then
                    RValue.Value = My.Resources.Speech2
                Else
                    RValue.Value = My.Resources.speech
                End If
            End If

            If LValue.Tag <> "Done" Then
                If StaffQueryGrid.Item("Light", e.RowIndex).Value >= 1 Then
                    LValue.Value = My.Resources.Red
                ElseIf StaffQueryGrid.Item("Light", e.RowIndex).Value = 1 Then
                    LValue.Value = My.Resources.Amber
                ElseIf StaffQueryGrid.Item("Light", e.RowIndex).Value <= 1 Then
                    LValue.Value = My.Resources.Green
                End If
                RValue.Tag = "Done"
                LValue.Tag = "Done"
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick

        Call Form1.UpdateCounter(Label8)

    End Sub


End Class