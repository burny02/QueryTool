Imports System
Imports System.Reflection
Imports System.Windows.Forms
Public Class Form1


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Visible = False
        Me.Hide()

        Call StartUpCentral()

        If AccessLevel = 0 Then

            RespView = New ResponseView
            Specifics(RespView)

            RespView.FilterCombo1.LiveData = False
            RespView.FilterCombo1.SetAsExternalSource("SiteCode", "Site", "SELECT DISTINCT Code AS SiteCode, Site FROM SiteCode", Overclass)
            RespView.FilterCombo2.LiveData = False
            RespView.FilterCombo2.SetAsExternalSource("RespondCode", "Group", "SELECT DISTINCT Code AS RespondCode, Group FROM GroupCode", Overclass)
            RespView.StaffQueryGrid.Columns("QueryID").Visible = False

            RespView.Text = SolutionName
            RespView.ShowDialog()
            Application.Exit()

        Else

            Me.Visible = True
            Me.WindowState = FormWindowState.Maximized

        End If


        Try
            Me.Label2.Text = "Query Tool " & vbNewLine & "Developed by David Burnside" & vbNewLine & "Version:     " & System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
        Catch ex As Exception
            Me.Label2.Text = "Query Tool " & vbNewLine & "Developed by David Burnside"
        End Try

        Me.Text = SolutionName

    End Sub

    Private Sub TabControl1_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabControl1.Selecting

        Dim SQLCode As String = vbNullString
        Dim Bind As BindingSource = BindingSource1
        Dim ctl As Object = Nothing

        If Overclass.UnloadData() = True Then
            e.Cancel = True
            Exit Sub
        End If

        Overclass.ResetCollection()
        RespondCommands.Clear()

        Select Case e.TabPage.Text


            Case "Add Queries"
                NewQueryGrid.Columns.Clear()
                Specifics(NewQueryGrid)

            Case "Reports"
                Me.DateTimePicker2.Value = Date.Now

        End Select


        Call Specifics(ctl)

    End Sub


    Public Sub Specifics(ctl As Object)

        If IsNothing(ctl) Then Exit Sub

        Dim SqlCode As String = vbNullString

        Select Case ctl.name

            Case "ResponseView"

                Dim CurrentFilter As String = ""
                Dim Filter1 As String = ""
                Dim Filter2 As String = ""
                Dim Filter3 As String = ""
                Dim Filter4 As String = ""
                Dim Filter5 As String = ""
                Dim Filter6 As String = ""
                Dim Filter7 As String = ""

                Try
                    CurrentFilter = (Overclass.CurrentDataSet.Tables(0).DefaultView.RowFilter)
                Catch ex As Exception
                End Try

                Try
                    Filter1 = RespView.FilterCombo1.Text
                Catch ex As Exception
                End Try

                Try
                    Filter2 = RespView.FilterCombo2.Text
                Catch ex As Exception
                End Try

                Try
                    Filter3 = RespView.FilterCombo3.Text
                Catch ex As Exception
                End Try

                Try
                    Filter4 = RespView.FilterCombo30.Text
                Catch ex As Exception
                End Try

                Try
                    Filter5 = RespView.FilterCombo90.Text
                Catch ex As Exception
                End Try

                Try
                    Filter6 = RespView.FilterCombo5.Text
                Catch ex As Exception
                End Try

                Try
                    Filter7 = RespView.FilterCombo4.Text
                Catch ex As Exception
                End Try

                RespView.StaffQueryGrid.Columns.Clear()
                SqlCode = "SELECT QName, CreatedByRole, QueryID, Study, Person, " &
                "Priority, Initials & ' ' & RVLID AS Volunteer, " &
                "VisitName, FormName, PageNo, Description, SiteCode, RespondCode, Bounced " &
                "FROM Queries WHERE Status ='Open' ORDER BY CreateDate ASC"
                Overclass.CreateDataSet(SqlCode, RespView.BindingSource1, RespView.StaffQueryGrid)

                With RespView.StaffQueryGrid
                    .ReadOnly = True
                    .Columns("QueryID").Visible = False
                    .Columns("Bounced").Visible = False
                    .Columns("Study").Visible = False
                    .Columns("SiteCode").Visible = False
                    .Columns("RespondCode").Visible = False
                    .Columns("Priority").Visible = False
                    .Columns("VisitName").HeaderText = "Study Visit"
                    .Columns("FormName").HeaderText = "Assessment/Procedure"
                    .Columns("PageNo").HeaderText = "Page No"
                    .Columns("Person").HeaderText = "Assigned"
                    .Columns("CreatedByRole").HeaderText = "Raised By"
                    .Columns("QName").HeaderText = "Cohort"
                    Dim clm2 As New DataGridViewImageColumn
                    clm2.HeaderText = "Respond"
                    clm2.Name = "RespondClm"
                    clm2.ImageLayout = DataGridViewImageCellLayout.Zoom
                    clm2.Image = My.Resources.speech
                    .Columns.Add(clm2)
                    .Columns("QName").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                    .Columns("RespondClm").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                    .Columns("PageNo").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                    .Columns("Person").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                    .Columns("Volunteer").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                    .Columns("CreatedByRole").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                End With

                With RespView
                    .FilterCombo5.SetAsInternalSource("QName", "QName", Overclass)
                    .FilterCombo4.SetAsInternalSource("CreatedByRole", "CreatedByRole", Overclass)
                    .FilterCombo3.AllowBlanks = False
                    .FilterCombo3.SetAsInternalSource("Study", "Study", Overclass)
                    .FilterCombo30.SetAsInternalSource("Priority", "Priority", Overclass)
                    .FilterCombo90.SetAsInternalSource("Volunteer", "Volunteer", Overclass)

                    .StaffQueryGrid.Columns("QueryID").Visible = False
                End With

                If CurrentFilter <> "" Then Overclass.CurrentDataSet.Tables(0).DefaultView.RowFilter = CurrentFilter
                RespView.FilterCombo1.Text = Filter1
                RespView.FilterCombo2.Text = Filter2
                RespView.FilterCombo3.Text = Filter3
                RespView.FilterCombo30.Text = Filter4
                RespView.FilterCombo90.Text = Filter5
                RespView.FilterCombo5.Text = Filter6
                RespView.FilterCombo4.Text = Filter7

                Try
                    RespView.FilterCombo1.RefreshCombo()
                Catch ex As Exception
                End Try
                Try
                    RespView.FilterCombo2.RefreshCombo()
                Catch ex As Exception
                End Try
                Try
                    RespView.FilterCombo3.RefreshCombo()
                Catch ex As Exception
                End Try
                Try
                    RespView.FilterCombo30.RefreshCombo()
                Catch ex As Exception
                End Try
                Try
                    RespView.FilterCombo4.RefreshCombo()
                Catch ex As Exception
                End Try
                Try
                    RespView.FilterCombo90.RefreshCombo()
                Catch ex As Exception
                End Try
                Try
                    RespView.FilterCombo5.RefreshCombo()
                Catch ex As Exception
                End Try


            Case "NewQueryGrid"

                NewQueryGrid.Columns.Clear()

                SqlCode = "SELECT QName, CreatedByRole, SiteCode, RespondCode, Person, TypeCode, Status, Study, QueryID, " &
                    "ClosedDate, ClosedBy, ClosedByRole, RVLID, Initials, " &
                    "VisitName, FormName, PageNo, Description, Priority, Bounced, AssCode, PDFLink " &
                    "FROM Queries " &
                    "WHERE Status='Open' OR Status='Responded' " &
                    "ORDER BY Status DESC, RVLID ASC"

                NewQueryGrid.AutoGenerateColumns = True
                Overclass.CreateDataSet(SqlCode, BindingSource1, NewQueryGrid)
                NewQueryGrid.AutoGenerateColumns = False

                FilterCombo2.SetAsInternalSource("QName", "QName", Overclass)
                FilterCombo1.SetAsInternalSource("CreatedByRole", "CreatedByRole", Overclass)
                FilterCombo7.LiveData = False
                FilterCombo7.SetAsExternalSource("SiteCode", "Site", "SELECT DISTINCT Code AS SiteCode, Site FROM SiteCode", Overclass)
                FilterCombo30.LiveData = False
                FilterCombo30.AllowBlanks = False
                FilterCombo30.SetAsExternalSource("Study", "Study", "SELECT StudyCode As Study FROM Study " &
                "WHERE Hidden=False ORDER BY StudyCode ASC", Overclass)
                FilterCombo30.SetDGVDefault(ctl, "Study")

                FilterCombo90.SetAsInternalSource("Initials", "Initials", Overclass)
                FilterCombo100.SetAsInternalSource("Status", "Status", Overclass)

                FilterCombo20.SetAsInternalSource("RVLID", "RVLID", Overclass)
                FilterCombo10.SetAsInternalSource("VisitName", "VisitName", Overclass)

                NewQueryGrid.Columns("Bounced").Visible = False
                NewQueryGrid.Columns("QueryID").Visible = False
                NewQueryGrid.Columns("Study").Visible = False
                NewQueryGrid.Columns("ClosedDate").Visible = False
                NewQueryGrid.Columns("ClosedBy").Visible = False
                NewQueryGrid.Columns("ClosedByRole").Visible = False
                NewQueryGrid.Columns("Study").Visible = False
                NewQueryGrid.Columns("Status").Visible = False
                NewQueryGrid.Columns("AssCode").Visible = False
                NewQueryGrid.Columns("PDFLink").Visible = False
                NewQueryGrid.Columns("QName").Visible = False

                NewQueryGrid.Columns("RVLID").HeaderText = "Subject"
                NewQueryGrid.Columns("VisitName").HeaderText = "Study Visit"
                NewQueryGrid.Columns("FormName").HeaderText = "Assessment/Procedure"
                NewQueryGrid.Columns("PageNo").HeaderText = "Page"
                NewQueryGrid.Columns("Description").HeaderText = "Query Description"
                NewQueryGrid.Columns("CreatedByRole").HeaderText = "Role"

                Dim DTArray(3) As DataTable
                Dim SqlArray(3) As String
                SqlArray(0) = "SELECT Site AS Display, Code FROM SiteCode WHERE Hidden=False ORDER BY Site"
                SqlArray(1) = "SELECT Group as Display, Code FROM GroupCode WHERE Hidden=False ORDER BY Group"
                SqlArray(2) = "SELECT ErrorType as Display, Code FROM TypeCode WHERE Hidden=False ORDER BY ErrorType"
                SqlArray(3) = "SELECT AssName, AssCode From AssType WHERE Hidden=false ORDER BY AssName"
                DTArray = Overclass.MultiTempDataTable(SqlArray)

                Dim clm3 As New DataGridViewComboBoxColumn
                clm3.DataSource = DTArray(0)
                clm3.DataPropertyName = "SiteCode"
                clm3.DisplayMember = "Display"
                clm3.ValueMember = "Code"
                clm3.Name = "SiteClm"
                clm3.HeaderText = "Site"
                NewQueryGrid.Columns.Add(clm3)

                Dim clm4 As New DataGridViewComboBoxColumn
                clm4.DataSource = DTArray(1)
                clm4.DataPropertyName = "RespondCode"
                clm4.DisplayMember = "Display"
                clm4.ValueMember = "Code"
                clm4.Name = "GroupClm"
                clm4.HeaderText = "Group"
                NewQueryGrid.Columns.Add(clm4)

                Dim clm5 As New DataGridViewComboBoxColumn
                clm5.DataSource = DTArray(2)
                clm5.DataPropertyName = "TypeCode"
                clm5.DisplayMember = "Display"
                clm5.ValueMember = "Code"
                clm5.Name = "TypeClm"
                clm5.HeaderText = "Type"
                NewQueryGrid.Columns.Add(clm5)

                Dim clm6 As New DataGridViewComboBoxColumn
                clm6.DataSource = DTArray(3)
                clm6.DisplayMember = "AssName"
                clm6.ValueMember = "AssCode"
                clm6.DataPropertyName = "AssCode"
                clm6.Name = "AssDrop"
                clm6.HeaderText = "Assessment Type"


                Dim clm7 As DataGridViewColumn = Overclass.SetUpNewComboColumn("SELECT QName AS Display FROM StudyCohort " &
                                                   "a inner join Study b ON a.CodeList=b.CodeList " &
                                                   "WHERE CStr(StudyCode)=", FilterCombo30,
                                                  "Display", "Display", "QName", "Cohort", NewQueryGrid, "CohortClm")

                clm7.SortMode =
                NewQueryGrid.Columns.Add(clm6)

                Dim cmb2 As New DataGridViewImageColumn
                cmb2.HeaderText = ""
                cmb2.Image = My.Resources.copy
                cmb2.ImageLayout = DataGridViewImageCellLayout.Zoom
                cmb2.Name = "CopyQuery"

                NewQueryGrid.Columns.Add(cmb2)


                Dim clm2 As New DataGridViewImageColumn
                clm2.HeaderText = ""
                clm2.Name = "RespondClm"
                clm2.ImageLayout = DataGridViewImageCellLayout.Zoom
                clm2.Image = My.Resources.speech
                NewQueryGrid.Columns.Add(clm2)


                Dim cmb As New DataGridViewImageColumn
                cmb.HeaderText = ""
                cmb.Image = My.Resources.TICK
                cmb.ImageLayout = DataGridViewImageCellLayout.Zoom
                cmb.Name = "StatusCmb"

                NewQueryGrid.Columns.Add(cmb)

                Dim pdfClm As New DataGridViewImageColumn
                pdfClm.HeaderText = ""
                pdfClm.Name = "PDF"
                pdfClm.ImageLayout = DataGridViewImageCellLayout.Zoom
                pdfClm.Image = My.Resources.PDF
                NewQueryGrid.Columns.Add(pdfClm)


                NewQueryGrid.Columns("SiteCode").Visible = False
                NewQueryGrid.Columns("RespondCode").Visible = False
                NewQueryGrid.Columns("TypeCode").Visible = False

                'Visible
                NewQueryGrid.Columns("CreatedByRole").DisplayIndex = 0
                NewQueryGrid.Columns("RVLID").DisplayIndex = 1
                NewQueryGrid.Columns("Initials").DisplayIndex = 2
                NewQueryGrid.Columns("CohortClm").DisplayIndex = 3
                NewQueryGrid.Columns("VisitName").DisplayIndex = 4
                NewQueryGrid.Columns("AssDrop").DisplayIndex = 5
                NewQueryGrid.Columns("FormName").DisplayIndex = 6
                NewQueryGrid.Columns("PageNo").DisplayIndex = 7
                NewQueryGrid.Columns("Description").DisplayIndex = 8
                NewQueryGrid.Columns("Priority").DisplayIndex = 9
                NewQueryGrid.Columns("SiteClm").DisplayIndex = 10
                NewQueryGrid.Columns("TypeClm").DisplayIndex = 11
                NewQueryGrid.Columns("Person").DisplayIndex = 12
                NewQueryGrid.Columns("GroupClm").DisplayIndex = 13
                NewQueryGrid.Columns("PDF").DisplayIndex = 14
                NewQueryGrid.Columns("CopyQuery").DisplayIndex = 15
                NewQueryGrid.Columns("RespondClm").DisplayIndex = 16
                NewQueryGrid.Columns("StatusCmb").DisplayIndex = 17


                'Invisible
                NewQueryGrid.Columns("SiteCode").DisplayIndex = 18
                NewQueryGrid.Columns("RespondCode").DisplayIndex = 19
                NewQueryGrid.Columns("TypeCode").DisplayIndex = 20
                NewQueryGrid.Columns("Status").DisplayIndex = 21
                NewQueryGrid.Columns("Study").DisplayIndex = 22
                NewQueryGrid.Columns("QueryID").DisplayIndex = 23
                NewQueryGrid.Columns("ClosedDate").DisplayIndex = 24
                NewQueryGrid.Columns("ClosedBy").DisplayIndex = 25
                NewQueryGrid.Columns("ClosedByRole").DisplayIndex = 26
                NewQueryGrid.Columns("Bounced").DisplayIndex = 27


                'Widths
                NewQueryGrid.Columns("RVLID").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                NewQueryGrid.Columns("Initials").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                NewQueryGrid.Columns("PageNo").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                NewQueryGrid.Columns("Priority").AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                NewQueryGrid.Columns("CohortClm").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                NewQueryGrid.Columns("CreatedByRole").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                NewQueryGrid.Columns("SiteClm").AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                NewQueryGrid.Columns("TypeClm").AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                NewQueryGrid.Columns("Person").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                NewQueryGrid.Columns("GroupClm").AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                NewQueryGrid.Columns("AssDrop").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                NewQueryGrid.Columns("VisitName").AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                NewQueryGrid.Columns("FormName").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader

                NewQueryGrid.Columns("PDF").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                NewQueryGrid.Columns("CopyQuery").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                NewQueryGrid.Columns("RespondClm").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                NewQueryGrid.Columns("StatusCmb").AutoSizeMode = DataGridViewAutoSizeColumnMode.None

                NewQueryGrid.Columns("CreatedByRole").Width = 45
                NewQueryGrid.Columns("RVLID").Width = 60
                NewQueryGrid.Columns("Initials").Width = 40
                NewQueryGrid.Columns("Person").Width = 45

                NewQueryGrid.Columns("PDF").Width = 30
                NewQueryGrid.Columns("CopyQuery").Width = 30
                NewQueryGrid.Columns("RespondClm").Width = 30
                NewQueryGrid.Columns("StatusCmb").Width = 30

                NewQueryGrid.Columns("TypeClm").DefaultCellStyle.WrapMode = DataGridViewTriState.True

        End Select


    End Sub


    Private Sub NewQueryGrid_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles NewQueryGrid.CellDoubleClick

        If e.RowIndex < 0 Then Exit Sub

        If IsDBNull(NewQueryGrid.Item("QueryID", e.RowIndex).Value) Then
            MsgBox("Cannot action an unsaved query")
            Exit Sub
        End If

        If Role <> NewQueryGrid.Item("CreatedByRole", e.RowIndex).Value Then
            MsgBox("Cannot action this query due to role")
            Exit Sub
        End If

        If e.ColumnIndex = sender.columns("PDF").index Then
            If NewQueryGrid.Item("Status", e.RowIndex).Value <> "Responded" Then
                MsgBox("Query must be responded to link to PDF")
                Exit Sub
            End If
            'PDF Stuff
            Dim FilePath As String = ""
            Try
                FilePath = NewQueryGrid.Item("PDFLink", e.RowIndex).Value
            Catch ex As Exception
            End Try
            If FilePath = "" Then
                Dim fd As OpenFileDialog = New OpenFileDialog()

                fd.Title = "Open File Dialog"
                fd.InitialDirectory = "C:\"
                fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
                fd.FilterIndex = 2
                fd.RestoreDirectory = True
                fd.AutoUpgradeEnabled = False

                If fd.ShowDialog() <> Windows.Forms.DialogResult.OK Then
                    fd = Nothing
                    Exit Sub
                End If

                NewQueryGrid.Item("PDFLink", e.RowIndex).Value = fd.FileName
                NewQueryGrid.Rows(e.RowIndex).Tag = ""

                fd = Nothing

            Else
                If MsgBox("A file is already attached, do you want to remove it?", vbYesNo) = vbNo Then
                    Try
                        System.Diagnostics.Process.Start(FilePath)
                    Catch ex As Exception
                        MsgBox("Unable to find the file.")
                    End Try
                Else

                    NewQueryGrid.Item("PDFLink", e.RowIndex).Value = ""
                    NewQueryGrid.Rows(e.RowIndex).Tag = ""

                End If
            End If
        End If
        If e.ColumnIndex = sender.columns("CopyQuery").index Then

            If MsgBox("Do you want to copy this query to a new line?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                Dim NewRow As DataRow = Overclass.CurrentDataSet.Tables(0).NewRow
                NewRow.Item("VisitName") = NewQueryGrid.Item("VisitName", e.RowIndex).Value
                NewRow.Item("FormName") = NewQueryGrid.Item("FormName", e.RowIndex).Value
                NewRow.Item("PageNo") = NewQueryGrid.Item("PageNo", e.RowIndex).Value
                NewRow.Item("Description") = NewQueryGrid.Item("Description", e.RowIndex).Value
                NewRow.Item("Priority") = NewQueryGrid.Item("Priority", e.RowIndex).Value
                NewRow.Item("QName") = NewQueryGrid.Item("QName", e.RowIndex).Value
                NewRow.Item("Study") = Trim(NewQueryGrid.Item("Study", e.RowIndex).Value)
                NewRow.Item("RVLID") = NewQueryGrid.Item("RVLID", e.RowIndex).Value
                NewRow.Item("Initials") = NewQueryGrid.Item("Initials", e.RowIndex).Value
                NewRow.Item("AssCode") = NewQueryGrid.Item("AssCode", e.RowIndex).Value
                NewRow.Item("SiteCode") = Trim(NewQueryGrid.Item("SiteCode", e.RowIndex).Value)
                NewRow.Item("TypeCode") = Trim(NewQueryGrid.Item("TypeCode", e.RowIndex).Value)
                NewRow.Item("Person") = NewQueryGrid.Item("Person", e.RowIndex).Value
                NewRow.Item("RespondCode") = NewQueryGrid.Item("RespondCode", e.RowIndex).Value
                NewRow.Item("Status") = "Open"


                Overclass.CurrentDataSet.Tables(0).Rows.Add(NewRow)
                NewQueryGrid.CurrentCell = NewQueryGrid.Item("RVLID", NewQueryGrid.NewRowIndex)

            End If
        End If

        If e.ColumnIndex = sender.columns("StatusCmb").index Then

            If Me.NewQueryGrid.Item("Status", e.RowIndex).Value = "Closed" Then Exit Sub
            If IsDBNull(NewQueryGrid.Item("QueryID", e.RowIndex).Value) = True Then Exit Sub

            If MsgBox("Are you sure you want to close this query?", vbYesNo) = vbYes Then
                Me.NewQueryGrid.Item("ClosedDate", e.RowIndex).Value = Format(DateTime.Now, "dd-MMM-yyyy HH:mm")
                Me.NewQueryGrid.Item("ClosedBy", e.RowIndex).Value = Overclass.GetUserName
                Me.NewQueryGrid.Item("ClosedByRole", e.RowIndex).Value = Role
                Me.NewQueryGrid.Item("Status", e.RowIndex).Value = "Closed"
                NewQueryGrid.Rows(e.RowIndex).Tag = ""
                BindingSource1.EndEdit()
            End If
        End If

        If e.ColumnIndex = sender.columns("RespondClm").index Then

            If NewQueryGrid.Item("Status", e.RowIndex).Value <> "Responded" Then Exit Sub

            Dim RespondText As String = vbNullString

            Dim QueryID As String
            QueryID = NewQueryGrid.Item("QueryID", e.RowIndex).Value

            If IsDBNull(NewQueryGrid.Item("QueryID", e.RowIndex).Value) = False Then
                Dim CSVString As String = Overclass.CreateCSVString(
                "Select format(Response_Timestamp,'dd-MMM-yyyy HH:mm') & ' (' & response_Person & ')  -  ' " &
                "& replace(response_text,',',';') FROM Response WHERE QueryID=" & QueryID)
                CSVString = Replace(CSVString, ",", vbNewLine & vbNewLine)
                If CSVString = "" Then CSVString = "No history found"
                RespondText = CSVString
            End If

            RespondText = RespondText & vbNewLine & "Please input response to query:"

            Dim Response = InputBox(RespondText, "Query Response")
            If Response = "" Then
                Exit Sub
            Else
                Dim SQL As String
                'NewQueryGrid.Item("Status", e.RowIndex).Value = "Open"
                Dim foundRows() As DataRow
                foundRows = Overclass.CurrentDataSet.Tables(0).Select("QueryID=" & QueryID)
                foundRows(0).Item("Bounced") = "True"
                foundRows(0).Item("Status") = "Open"
                foundRows(0).EndEdit()
                SQL = "INSERT INTO Response(QueryID,Response_Text,Response_Person) " &
                "VALUES (" & QueryID & ", '" & Response & "', '" & Overclass.GetUserName & "')"

                Dim Cmd As OleDb.OleDbCommand
                Cmd = New OleDb.OleDbCommand(SQL)
                Overclass.SetCommandConnection(Cmd)
                RespondCommands.Add(Cmd)
                NewQueryGrid.Rows(e.RowIndex).Tag = ""
            End If
        End If

    End Sub

    Private Sub NewQueryGrid_DefaultValuesNeeded_1(sender As Object, e As DataGridViewRowEventArgs) Handles NewQueryGrid.DefaultValuesNeeded
        NewQueryGrid.Item("CreatedByRole", e.Row.Index).Value = Role
        NewQueryGrid.Item("Status", e.Row.Index).Value = "Open"
    End Sub

    Private Sub NewQueryGrid_DataError(sender As Object, e As DataGridViewDataErrorEventArgs)

    End Sub

    Private Sub NewQueryGrid_RowPrePaint(sender As Object, e As DataGridViewRowPrePaintEventArgs) Handles NewQueryGrid.RowPrePaint

        Try

            If NewQueryGrid.Rows(e.RowIndex).Tag = "" Then

                Dim FilePath As String = ""
                Try
                    FilePath = NewQueryGrid.Item("PDFLink", e.RowIndex).Value
                Catch ex As Exception
                End Try
                If FilePath = "" Then
                    NewQueryGrid.Item("PDF", e.RowIndex).Value = My.Resources.EmptyFile
                Else
                    NewQueryGrid.Item("PDF", e.RowIndex).Value = My.Resources.PDF
                End If


                If IsDBNull(NewQueryGrid.Item("QueryID", e.RowIndex).Value) Then
                    NewQueryGrid.Item("StatusCmb", e.RowIndex).Value = My.Resources.hyphen
                    NewQueryGrid.Item("RespondClm", e.RowIndex).Value = My.Resources.hyphen
                    NewQueryGrid.Item("PDF", e.RowIndex).Value = My.Resources.hyphen
                    NewQueryGrid.Item("CopyQuery", e.RowIndex).Value = My.Resources.hyphen
                End If

                If Role <> NewQueryGrid.Item("CreatedByRole", e.RowIndex).Value Then
                    NewQueryGrid.Item("StatusCmb", e.RowIndex).Value = My.Resources.hyphen
                    NewQueryGrid.Item("RespondClm", e.RowIndex).Value = My.Resources.hyphen
                    NewQueryGrid.Item("PDF", e.RowIndex).Value = My.Resources.hyphen
                    NewQueryGrid.Item("CopyQuery", e.RowIndex).Value = My.Resources.hyphen
                    If NewQueryGrid.Item("CreatedByRole", e.RowIndex).Value <> "" Then NewQueryGrid.Rows(e.RowIndex).ReadOnly = True
                End If

                If NewQueryGrid.Item("Status", e.RowIndex).Value = "Closed" Then NewQueryGrid.Item("StatusCmb", e.RowIndex).Value = My.Resources.hyphen

                If NewQueryGrid.Item("Status", e.RowIndex).Value <> "Responded" Then
                    NewQueryGrid.Item("RespondClm", e.RowIndex).Value = My.Resources.hyphen
                    NewQueryGrid.Item("PDF", e.RowIndex).Value = My.Resources.hyphen
                End If



                NewQueryGrid.Item("CreatedByRole", e.RowIndex).ReadOnly = True
                NewQueryGrid.Rows(e.RowIndex).Tag = "DontPaint"



            End If

        Catch ex As Exception
        End Try

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        Call UpdateCounter(Label6)

    End Sub


    Public Sub UpdateCounter(WhatLabel As Label)

        Try
            WhatLabel.Text = "Displaying " & Overclass.CurrentDataSet.Tables(0).DefaultView.Count & " queries"
        Catch ex As Exception
        End Try

    End Sub

End Class