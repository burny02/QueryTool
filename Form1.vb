Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Visible = False
        Me.Hide()

        Call StartUpCentral()

        If AccessLevel = 1 Then

            Dim AddQry As New AddQuery
            AddQry.ShowDialog()

        Else

            Me.Visible = True
            Me.WindowState = FormWindowState.Maximized

        End If

        
        Try
            Me.Label2.Text = "Query Tool " & vbNewLine & "Developed by David Burnside" & vbNewLine & "Version: " & System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
        Catch ex As Exception
            Me.Label2.Text = "Query Tool " & vbNewLine & "Developed by David Burnside"
        End Try

        Me.Text = SolutionName

    End Sub

    Private Sub TabControl1_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabControl1.Selecting

        Dim SQLCode As String = vbNullString
        Dim Bind As BindingSource = BindingSource1
        Dim ctl As Object = Nothing

        If overclass.UnloadData() = True Then
            e.Cancel = True
            Exit Sub
        End If

        Overclass.ResetCollection()

        Select Case e.TabPageIndex

            Case 1
                ctl = Me.DataGridView1
                SQLCode = "SELECT DisplayName, UploadDate, UploadPerson FROM Study ORDER BY UploadDate DESC"
                overclass.CreateDataSet(SQLCode, Bind, ctl)

            Case 2
                StartCombo(Me.ComboBox1)
                
            Case 3
                StartCombo(Me.ComboBox2)
                

            Case 4
                StartCombo(Me.ComboBox3)

            Case 5
                Me.DateTimePicker2.Value = Date.Now

        End Select


        Call Specifics(ctl)

    End Sub

    Private Sub Grid2And3(ctl As Object, Combo As ComboBox, SQLString As String)

        'Call ResetDataGrid()
        Overclass.CreateDataSet(SQLString, BindingSource1, ctl)
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
        cmb.DataSource = Overclass.TempDataTable("SELECT Code & ' - ' & Site AS Display, Code FROM SiteCode a inner join Study b ON a.ListID=b.CodeList " & _
                                     "WHERE StudyCode='" & Combo.SelectedValue.ToString & "' ORDER BY Code ASC")
        cmb.DataPropertyName = Overclass.CurrentDataSet.Tables(0).Columns(1).ToString
        cmb.ValueMember = "Code"
        cmb.DisplayMember = "Display"
        ctl.Columns.Add(cmb)
        Dim cmb2 As New DataGridViewComboBoxColumn()
        cmb2.DataSource = Overclass.TempDataTable("SELECT Code & ' - ' & ErrorType AS Display, Code FROM TypeCode a inner join Study b ON a.ListID=b.CodeList " & _
                                     "WHERE StudyCode='" & Combo.SelectedValue.ToString & "' ORDER BY Code ASC")
        cmb2.DataPropertyName = Overclass.CurrentDataSet.Tables(0).Columns(2).ToString
        cmb2.ValueMember = "Code"
        cmb2.DisplayMember = "Display"
        ctl.Columns.Add(cmb2)
        Dim cmb3 As New DataGridViewComboBoxColumn()
        cmb3.DataSource = Overclass.TempDataTable("SELECT Code & ' - ' & Group AS Display, Code FROM GroupCode a inner join Study b ON a.ListID=b.CodeList " & _
                                     "WHERE StudyCode='" & Combo.SelectedValue.ToString & "' ORDER BY Code ASC")
        cmb3.DataPropertyName = Overclass.CurrentDataSet.Tables(0).Columns(4).ToString
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

    Public Sub Specifics(ctl As Object)

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

                Dim AllowedSite As String = Overclass.CreateCSVString("SELECT Code FROM SiteCODE a INNER JOIN Study b ON a.ListID=b.CodeList " & _
                                                            "WHERE StudyCode='" & Me.ComboBox2.SelectedValue.ToString & "'")
                Dim AllowedResponse As String = Overclass.CreateCSVString("SELECT Code FROM GroupCode a INNER JOIN Study b ON a.ListID=b.CodeList " & _
                                                            "WHERE StudyCode='" & Me.ComboBox2.SelectedValue.ToString & "'")
                Dim AllowedType As String = Overclass.CreateCSVString("SELECT Code FROM TypeCode a INNER JOIN Study b ON a.ListID=b.CodeList " & _
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

End Class