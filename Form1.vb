Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

        Call StartUpCentral()

        Call Central.LockCheck()
        Call Central.LoginCheck()

        Try
            Me.Label2.Text = "Query Tool " & vbNewLine & "Developed by David Burnside" & vbNewLine & "Version: " & System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
        Catch ex As Exception
            Me.Label2.Text = "Query Tool " & vbNewLine & "Developed by David Burnside"
        End Try

        Me.Text = SolutionName

    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        If Central.UnloadData() = True Then e.Cancel = True
        Call Central.Quitter(True)

    End Sub

    Private Sub TabControl1_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabControl1.Selecting

        Dim SQLCode As String = vbNullString
        Dim Bind As BindingSource = BindingSource1
        Dim ctl As Object = Nothing

        If Central.UnloadData() = True Then
            e.Cancel = True
            Exit Sub
        End If

        Call ResetDataGrid()

        Select Case e.TabPageIndex

            Case 1
                ctl = Me.DataGridView1
                SQLCode = "SELECT DisplayName, UploadDate, UploadPerson FROM Study ORDER BY UploadDate DESC"
                Central.CreateDataSet(SQLCode, Bind, ctl)

            Case 2
                ctl = Me.DataGridView2
                Me.ComboBox1.DataSource = Central.TempDataSet("SELECT DisplayName, StudyCode FROM Study ORDER BY StudyCode ASC").Tables(0)
                Me.ComboBox1.ValueMember = "StudyCode"
                Me.ComboBox1.DisplayMember = "DisplayName"

            Case 3
                ctl = Me.DataGridView3
                Me.ComboBox2.DataSource = Central.TempDataSet("SELECT DisplayName, StudyCode FROM Study ORDER BY StudyCode ASC").Tables(0)
                Me.ComboBox2.ValueMember = "StudyCode"
                Me.ComboBox2.DisplayMember = "DisplayName"

            Case 4
                Me.ComboBox3.DataSource = Central.TempDataSet("SELECT DisplayName, StudyCode FROM Study ORDER BY StudyCode ASC").Tables(0)
                Me.ComboBox3.ValueMember = "StudyCode"
                Me.ComboBox3.DisplayMember = "DisplayName"

        End Select


        Call Specifics(ctl)

    End Sub

    Private Sub ResetDataGrid()

        Me.DataGridView1.Columns.Clear()
        Me.DataGridView1.DataSource = Nothing
        Me.DataGridView2.Columns.Clear()
        Me.DataGridView2.DataSource = Nothing
        Me.DataGridView3.Columns.Clear()
        Me.DataGridView3.DataSource = Nothing

    End Sub

    Private Sub Grid2And3(ctl As Object, Combo As ComboBox, SQLString As String)

        Call ResetDataGrid()
        Central.CreateDataSet(SQLString, BindingSource1, ctl)
        ctl.columns(0).visible = False
        ctl.columns(1).visible = False
        ctl.columns(2).visible = False
        ctl.columns(4).visible = False
        ctl.columns(5).readonly = True
        ctl.columns(6).readonly = True
        ctl.columns(7).readonly = True
        ctl.columns(8).readonly = True
        ctl.AllowUserToAddRows = False
        Dim cmb As New DataGridViewComboBoxColumn()
        cmb.DataSource = Central.TempDataSet("SELECT Code & ' - ' & Site AS Display, Code FROM SiteCode a inner join Study b ON a.ListID=b.CodeList " & _
                                     "WHERE StudyCode='" & Combo.SelectedValue.ToString & "' ORDER BY Code ASC").Tables(0)
        cmb.DataPropertyName = Central.CurrentDataSet.Tables(0).Columns(1).ToString
        cmb.ValueMember = "Code"
        cmb.DisplayMember = "Display"
        ctl.Columns.Add(cmb)
        Dim cmb2 As New DataGridViewComboBoxColumn()
        cmb2.DataSource = Central.TempDataSet("SELECT Code & ' - ' & ErrorType AS Display, Code FROM TypeCode a inner join Study b ON a.ListID=b.CodeList " & _
                                     "WHERE StudyCode='" & Combo.SelectedValue.ToString & "' ORDER BY Code ASC").Tables(0)
        cmb2.DataPropertyName = Central.CurrentDataSet.Tables(0).Columns(2).ToString
        cmb2.ValueMember = "Code"
        cmb2.DisplayMember = "Display"
        ctl.Columns.Add(cmb2)
        Dim cmb3 As New DataGridViewComboBoxColumn()
        cmb3.DataSource = Central.TempDataSet("SELECT Code & ' - ' & Group AS Display, Code FROM GroupCode a inner join Study b ON a.ListID=b.CodeList " & _
                                     "WHERE StudyCode='" & Combo.SelectedValue.ToString & "' ORDER BY Code ASC").Tables(0)
        cmb3.DataPropertyName = Central.CurrentDataSet.Tables(0).Columns(4).ToString
        cmb3.ValueMember = "Code"
        cmb3.DisplayMember = "Display"
        ctl.Columns.Add(cmb3)
        ctl.columns(3).displayindex = 10
        cmb3.HeaderText = "Respond Code"
        cmb2.HeaderText = "Type Code"
        cmb.HeaderText = "Site Code"
        ctl.columns(6).HeaderText = "Form Name"
        ctl.Columns(3).maxinputlength = 3

    End Sub

    Private Sub Specifics(ctl As Object)

        If IsNothing(ctl) Then Exit Sub

        Select Case ctl.name

            Case "DataGridView1"
                ctl.columns(0).headertext = "Study"
                ctl.columns(1).headertext = "Last Update"
                ctl.columns(2).headertext = "Upload Person"
                ctl.columns(1).DefaultCellStyle.Format = "dd-MMM-yyyy - HH:mm"
                ctl.enabled = False
                ctl.AllowUserToAddRows = False
            Case "DataGridView2"

                Dim SQLCode As String = "SELECT a.QueryID, SiteCode, TypeCode, Person, RespondCode, RVLID, " & _
                                "FormName, Description, Status FROM QueryCodes as a INNER JOIN Queries as b ON a.QueryID=b.QueryID " & _
                                "WHERE Study='" & Me.ComboBox1.SelectedValue.ToString & "' ORDER BY RVLID ASC"
                Call Grid2And3(ctl, Me.ComboBox1, SQLCode)

            Case "DataGridView3"

                Dim AllowedSite As String = Central.CreateCSVString("SELECT Code FROM SiteCODE a INNER JOIN Study b ON a.ListID=b.CodeList " & _
                                                            "WHERE StudyCode='" & Me.ComboBox2.SelectedValue.ToString & "'")
                Dim AllowedResponse As String = Central.CreateCSVString("SELECT Code FROM GroupCode a INNER JOIN Study b ON a.ListID=b.CodeList " & _
                                                            "WHERE StudyCode='" & Me.ComboBox2.SelectedValue.ToString & "'")
                Dim AllowedType As String = Central.CreateCSVString("SELECT Code FROM TypeCode a INNER JOIN Study b ON a.ListID=b.CodeList " & _
                                                            "WHERE StudyCode='" & Me.ComboBox2.SelectedValue.ToString & "'")

                Dim SQLCode As String = "SELECT a.QueryID, SiteCode, TypeCode, Person, RespondCode, RVLID, " & _
                                "FormName, Description, Status FROM QueryCodes as a INNER JOIN Queries as b ON a.QueryID=b.QueryID " & _
                                "WHERE Study='" & Me.ComboBox2.SelectedValue.ToString & "'" & _
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
                Call Grid2And3(ctl, Me.ComboBox2, SQLCode)

        End Select


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Call UploadCSV()

    End Sub

    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown

        e.SuppressKeyPress = True

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        If Me.ComboBox1.SelectedValue.ToString <> "System.Data.DataRowView" Then

            If Central.UnloadData() = True Then Exit Sub
            Call ResetDataGrid()
            Call Specifics(Me.DataGridView2)

        End If

    End Sub

    Private Sub DataGridView2_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView2.DataError

        Call Central.ErrorHandler(sender, e)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Call Saver(DataGridView2)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call Saver(DataGridView3)
    End Sub

    Private Sub DataGridView3_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView3.DataError
        Call Central.ErrorHandler(sender, e)
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If Me.ComboBox2.SelectedValue.ToString <> "System.Data.DataRowView" Then

            If Central.UnloadData() = True Then Exit Sub
            Call ResetDataGrid()
            Call Specifics(Me.DataGridView3)

        End If
    End Sub

    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
        e.SuppressKeyPress = True
    End Sub

    Private Sub Button6_Click_1(sender As Object, e As EventArgs) Handles Button6.Click
        Call ExportExcel("SELECT dateadd('d',QueryAgeLimit,CreateDate) AS DueDate," & _
                         "Person as [Allocated To], Site, Group, RVLID, " & _
                        "FormName, Description " & _
                        "FROM (((((Queries a INNER JOIN Study b ON a.Study=b.StudyCode) " & _
                        "INNER JOIN QueryCodes c ON a.QueryID=c.QueryID) " & _
                        "INNER JOIN GroupCode d ON b.CodeList=d.ListID) " & _
                        "INNER JOIN TypeCode e ON b.CodeList=e.ListID) " & _
                        "INNER JOIN SiteCode f ON b.CodeList=f.ListID) " & _
                        "WHERE Status='Open' AND Study='" & Me.ComboBox3.SelectedValue.ToString & "' " & _
                        " AND f.code=c.SiteCode AND c.RespondCode=d.code AND TypeCode=e.code" _
                         , Me.ComboBox3.SelectedValue.ToString, True)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Call ExportExcel("SELECT a.*, c.* " & _
                        "FROM (((((Queries a INNER JOIN Study b ON a.Study=b.StudyCode) " & _
                        "INNER JOIN QueryCodes c ON a.QueryID=c.QueryID) " & _
                        "INNER JOIN GroupCode d ON b.CodeList=d.ListID) " & _
                        "INNER JOIN TypeCode e ON b.CodeList=e.ListID) " & _
                        "INNER JOIN SiteCode f ON b.CodeList=f.ListID) " & _
                        "WHERE f.code=c.SiteCode AND c.RespondCode=d.code AND TypeCode=e.code" _
                         , Me.ComboBox3.SelectedValue.ToString, False)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call ExportExcel("SELECT " & _
                         "Person as [Allocated To], Site, Group, RVLID, " & _
                        "FormName, Description " & _
                        "FROM (((((Queries a INNER JOIN Study b ON a.Study=b.StudyCode) " & _
                        "INNER JOIN QueryCodes c ON a.QueryID=c.QueryID) " & _
                        "INNER JOIN GroupCode d ON b.CodeList=d.ListID) " & _
                        "INNER JOIN TypeCode e ON b.CodeList=e.ListID) " & _
                        "INNER JOIN SiteCode f ON b.CodeList=f.ListID) " & _
                        "WHERE Status='Responded' AND Study='" & Me.ComboBox3.SelectedValue.ToString & "' " & _
                        " AND f.code=c.SiteCode AND c.RespondCode=d.code AND TypeCode=e.code" _
                         , Me.ComboBox3.SelectedValue.ToString, False)
    End Sub

    Private Sub DataGridView3_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellEnter
        Call Central.SingleClick(sender, e)
    End Sub

    Private Sub DataGridView2_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellEnter
        Call Central.SingleClick(sender, e)
    End Sub
End Class