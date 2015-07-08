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
                SQLCode = "SELECT replace(StudyCode,'Retroscreen-',''), UploadDate, UploadPerson FROM Study ORDER BY UploadDate ASC"
                CreateDataSet(SQLCode, Bind, ctl)

            Case 2
                ctl = Me.DataGridView2
                

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
                ctl.columns(1).DefaultCellStyle.Format = "dd-MMM-yyyy"
                ctl.enabled = False
                ctl.AllowUserToAddRows = False
            Case "DataGridView2"
                Me.ComboBox1.DataSource = TempDataSet("SELECT replace(StudyCode,'Retroscreen-','') as Study1, StudyCode FROM Study ORDER BY StudyCode ASC").Tables(0)
                Me.ComboBox1.ValueMember = "StudyCode"
                Me.ComboBox1.DisplayMember = "Study1"
                Dim SQLCode As String = "SELECT a.*, Description, Status, FormName, RVLID FROM QueryCodes as a INNER JOIN Queries as b ON a.QueryID=b.QueryID " & _
                "WHERE Study='" & Me.ComboBox1.SelectedValue.ToString & "'"
                CreateDataSet(SQLCode, BindingSource1, ctl)
                ctl.columns(0).visible = False
                ctl.AllowUserToAddRows = False

        End Select


    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Call UploadCSV()

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

        Call UploadCSV()

    End Sub

    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown

        e.SuppressKeyPress = True

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        If Me.ComboBox1.SelectedValue.ToString <> "System.Data.DataRowView" Then

            Dim ctl As Object = Me.DataGridView2

            Call UnloadData()
            Dim SQLCode As String
            SQLCode = "SELECT a.*, Description, Status, FormName, RVLID FROM QueryCodes as a INNER JOIN Queries as b ON a.QueryID=b.QueryID " & _
                    "WHERE Study='" & Me.ComboBox1.SelectedValue.ToString & "'"
            CreateDataSet(SQLCode, BindingSource1, ctl)

        End If

    End Sub
End Class