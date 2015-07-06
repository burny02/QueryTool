Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call LockCheck()

        Call LoginCheck()

        Me.Label2.Text = "Developed by David Burnside" & vbNewLine & vbTab & "Version: " & My.Application.Info.Version.ToString()

        Me.Text = SolutionName

    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        If UnloadData() = True Then e.Cancel = True
        Call Quitter(True)

    End Sub

    Private Sub TabControl1_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabControl1.Selecting

        Dim SQLCode As String = vbNullString
        Dim Bind As BindingSource = BindingSource1
        Dim ctl As Object = Nothing

        If UnloadData() = True Then
            e.Cancel = True
            Exit Sub
        End If

        Call ResetDataGrid()

        Select Case e.TabPageIndex

            Case 1
                ctl = Me.DataGridView1
                SQLCode = "SELECT StudyCode, UploadDate, UploadPerson FROM Study ORDER BY UploadDate ASC"
                CreateDataSet(SQLCode, Bind, ctl)

        End Select


        Call Specifics(ctl)

    End Sub

    Private Sub ResetDataGrid()

        Me.DataGridView1.Columns.Clear()
        Me.DataGridView1.DataSource = Nothing

    End Sub

    Private Sub Specifics(ctl As Object)

        If IsNothing(ctl) Then Exit Sub

        Select Case ctl.name

            Case "DataGridView1"
                ctl.columns(0).headertext = "Study"
                ctl.columns(1).headertext = "Last Update"
                ctl.columns(2).headertext = "Upload Person"
                ctl.enabled = False
                ctl.AllowUserToAddRows = False
        End Select


    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Call UploadCSV()

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

        Call UploadCSV()

    End Sub
End Class