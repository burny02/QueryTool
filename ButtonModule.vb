Imports Microsoft.Reporting.WinForms

Module ButtonModule

    Public Sub ButtonSpecifics(sender As Object, e As EventArgs)

        Dim ctl As Object = Nothing

        Select Case sender.name.ToString

            Case "Button1"
                Call Saver(Form1.DataGridView2)

            Case "Button2"
                Call UploadCSV()


            Case "Button4"
                Call ExportExcel("SELECT " & _
                         "Person as [Allocated To], Site, Group, RVLID, " & _
                        "FormName, Description " & _
                        "FROM (((((Queries a INNER JOIN Study b ON a.Study=b.StudyCode) " & _
                        "INNER JOIN QueryCodes c ON a.QueryID=c.QueryID) " & _
                        "INNER JOIN GroupCode d ON b.CodeList=d.ListID) " & _
                        "INNER JOIN TypeCode e ON b.CodeList=e.ListID) " & _
                        "INNER JOIN SiteCode f ON b.CodeList=f.ListID) " & _
                        "WHERE Status='Responded' AND Study='" & Form1.ComboBox3.SelectedValue.ToString & "' " & _
                        " AND f.code=c.SiteCode AND c.RespondCode=d.code AND TypeCode=e.code" _
                         , Form1.ComboBox3.SelectedValue.ToString, False)

            Case "Button5"
                Call ExportExcel("SELECT a.*, c.* " &
                        "FROM (((((Queries a INNER JOIN Study b ON a.Study=b.StudyCode) " &
                        "INNER JOIN QueryCodes c ON a.QueryID=c.QueryID) " &
                        "INNER JOIN GroupCode d ON b.CodeList=d.ListID) " &
                        "INNER JOIN TypeCode e ON b.CodeList=e.ListID) " &
                        "INNER JOIN SiteCode f ON b.CodeList=f.ListID) " &
                        "WHERE f.code=c.SiteCode AND c.RespondCode=d.code AND TypeCode=e.code" &
                        " AND Study='" & Form1.ComboBox3.SelectedValue.ToString & "'", Form1.ComboBox3.SelectedValue.ToString, False)

            Case "Button6"
                Call ExportExcel("SELECT dateadd('d',QueryAgeLimit,CreateDate) AS DueDate," & _
                         "Person as [Allocated To], Site, Group, RVLID, " & _
                        "FormName, Description " & _
                        "FROM (((((Queries a INNER JOIN Study b ON a.Study=b.StudyCode) " & _
                        "INNER JOIN QueryCodes c ON a.QueryID=c.QueryID) " & _
                        "INNER JOIN GroupCode d ON b.CodeList=d.ListID) " & _
                        "INNER JOIN TypeCode e ON b.CodeList=e.ListID) " & _
                        "INNER JOIN SiteCode f ON b.CodeList=f.ListID) " & _
                        "WHERE Status='Open' AND Study='" & Form1.ComboBox3.SelectedValue.ToString & "' " & _
                        " AND f.code=c.SiteCode AND c.RespondCode=d.code AND TypeCode=e.code" _
                         , Form1.ComboBox3.SelectedValue.ToString, True)

            Case "Button7"
                CheckDates()
                Dim OK As New ReportViewer
                OK.Visible = True
                OK.ReportViewer1.Visible = True
                OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "Query_Management_Tool.AvgResponse.rdlc"
                OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet", _
                                                           Overclass.TempDataTable("SELECT * FROM AvgResponse " & _
                                                                                  "WHERE FilterDate Between " & Overclass.SQLDate(Form1.DateTimePicker1.Value) & _
                                                                                  " AND " & Overclass.SQLDate(Form1.DateTimePicker2.Value))))
                OK.ReportViewer1.RefreshReport()

            Case "Button8"
                CheckDates()
                Dim OK As New ReportViewer
                OK.Visible = True
                OK.ReportViewer1.Visible = True
                OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "Query_Management_Tool.Totals.rdlc"
                OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet",
                                                          Overclass.TempDataTable("SELECT * FROM Totals " &
                                                                                  "WHERE FilterDate Between " & Overclass.SQLDate(Form1.DateTimePicker1.Value) &
                                                                                  " AND " & Overclass.SQLDate(Form1.DateTimePicker2.Value))))
                OK.ReportViewer1.RefreshReport()

            Case "Button9"
                CheckDates()
                Dim OK As New ReportViewer
                OK.Visible = True
                OK.ReportViewer1.Visible = True
                OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "Query_Management_Tool.QCTeam.rdlc"
                OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet",
                                                          Overclass.TempDataTable("SELECT * FROM QCTeam " &
                                                                                  "WHERE FilterDate Between " & Overclass.SQLDate(Form1.DateTimePicker1.Value) &
                                                                                  " AND " & Overclass.SQLDate(Form1.DateTimePicker2.Value))))
                OK.ReportViewer1.RefreshReport()

            Case "Button10"
                CheckDates()
                Dim OK As New ReportViewer
                OK.Visible = True
                OK.ReportViewer1.Visible = True
                OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "Query_Management_Tool.Types.rdlc"
                OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet",
                                                          Overclass.TempDataTable("Select * FROM Types " &
                                                                                  "WHERE FilterDate Between " & Overclass.SQLDate(Form1.DateTimePicker1.Value) &
                                                                                  " And " & Overclass.SQLDate(Form1.DateTimePicker2.Value))))
                OK.ReportViewer1.RefreshReport()

            Case "Button11"
                CheckDates()
                Dim OK As New ReportViewer
                OK.Visible = True
                OK.ReportViewer1.Visible = True
                OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "Query_Management_Tool.Responders.rdlc"
                OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet",
                                                          Overclass.TempDataTable("Select * FROM Responders " &
                                                                                  "WHERE FilterDate Between " & Overclass.SQLDate(Form1.DateTimePicker1.Value) &
                                                                                  " And " & Overclass.SQLDate(Form1.DateTimePicker2.Value))))
                OK.ReportViewer1.RefreshReport()

            Case "Button12"
                CheckDates()
                Dim OK As New ReportViewer
                OK.Visible = True
                OK.ReportViewer1.Visible = True
                OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "Query_Management_Tool.QCIndividual.rdlc"
                OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet",
                                                          Overclass.TempDataTable("Select * FROM QCIndividual " &
                                                                                  "WHERE FilterDate Between " & Overclass.SQLDate(Form1.DateTimePicker1.Value) &
                                                                                  " And " & Overclass.SQLDate(Form1.DateTimePicker2.Value))))
                OK.ReportViewer1.RefreshReport()

            Case "Button13"
                CheckDates()
                Dim OK As New ReportViewer
                OK.Visible = True
                OK.ReportViewer1.Visible = True
                OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "Query_Management_Tool.ToolUsage.rdlc"
                OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet",
                                                          Overclass.TempDataTable("Select * FROM ToolUsage " &
                                                                                  "WHERE FilterDate Between " & Overclass.SQLDate(Form1.DateTimePicker1.Value) &
                                                                                  " And " & Overclass.SQLDate(Form1.DateTimePicker2.Value))))
                OK.ReportViewer1.RefreshReport()

            Case "Button14"
                CheckDates()
                Dim OK As New ReportViewer
                OK.Visible = True
                OK.ReportViewer1.Visible = True
                OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "Query_Management_Tool.Deviations.rdlc"
                OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet",
                                                          Overclass.TempDataTable("Select * FROM Deviations " &
                                                                                  "WHERE FilterDate Between " & Overclass.SQLDate(Form1.DateTimePicker1.Value) &
                                                                                  " And " & Overclass.SQLDate(Form1.DateTimePicker2.Value))))
                OK.ReportViewer1.RefreshReport()

            Case "Button15"
                CheckDates()
                Dim OK As New ReportViewer
                OK.Visible = True
                OK.ReportViewer1.Visible = True
                OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "Query_Management_Tool.DataClean.rdlc"
                OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet",
                                                          Overclass.TempDataTable("Select * FROM DataClean " &
                                                                                  "WHERE FilterDate Between " & Overclass.SQLDate(Form1.DateTimePicker1.Value) &
                                                                                  " And " & Overclass.SQLDate(Form1.DateTimePicker2.Value))))
                OK.ReportViewer1.RefreshReport()

            Case "Button16"

                AdQry = New AddQuery
                AddControls(AdQry)
                AdQry.Visible = True
                AdQry.TabControl1.Controls.Remove(AdQry.TabPage2)
                AdQry.TabControl1.Controls.Remove(AdQry.TabPage3)


            Case "Button17"

                Dim RoleCrit As String = vbNullString

                If MsgBox("Only correctly allocated queries will print out." & vbNewLine & vbNewLine &
                          "Do you want To print ONLY " & Role & " queries?" _
                          & vbNewLine & "Click NO For ALL open queries", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then

                    RoleCrit = " And CreatedByRole='" & Role & "'"

                End If


                Dim RVLID As Long = 0
                Dim InputString As String = vbNullString

                InputString = InputBox("Please input RVLID to print", "RVLID", "123456")

                If InputString = vbNullString Then Exit Sub

                Try
                    RVLID = CLng(InputString)
                Catch ex As Exception
                    Exit Sub
                End Try

                'SELECT RVLID FROM CURRENTDATA SET

                Try

                    Dim SqlString As String = "SELECT * FROM PrintOut WHERE RVLID='" & RVLID & "'" &
                                                                  " AND Status='Open'" & RoleCrit

                    Dim dt As DataTable = Overclass.TempDataTable(SqlString)

                    If dt.Rows.Count = 0 Then
                        MsgBox("No queries found for volunteer " & RVLID)
                        Exit Sub
                    End If

                    'RUN REPORT SEPERATING BY VISIT 

                    Dim OK As New ReportViewer
                    OK.Visible = True
                    OK.ReportViewer1.Visible = True
                    OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "Query_Management_Tool.PrintReport.rdlc"
                    OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet",
                                                               dt))
                    OK.ReportViewer1.RefreshReport()

                Catch ex As Exception

                    MsgBox(ex.Message)
                    Exit Sub

                End Try

            Case "Button101"
                Call Saver(AddQuery.NewQueryGrid)

            Case "Button102"
                Call Saver(AddQuery.NewQueryGrid2)

            Case "Button103"

                MsgBox("Only correctly allocated queries will print out. Please ensure queries are saved first")


                Dim RVLID As Long = 0
                Dim InputString As String = vbNullString

                InputString = InputBox("Please input RVLID to print", "RVLID", "123456")

                If InputString = vbNullString Then Exit Sub

                Try
                    RVLID = CLng(InputString)
                Catch ex As Exception
                    Exit Sub
                End Try

                'SELECT RVLID FROM CURRENTDATA SET

                Try

                    Dim SqlString As String = "SELECT * FROM PrintOut WHERE RVLID='" & RVLID & "'" &
                                                                  " AND CreatedByRole='" & Role & "'" &
                                                                  " AND Status='Open'"

                    Dim dt As DataTable = Overclass.TempDataTable(SqlString)

                    If dt.Rows.Count = 0 Then
                        MsgBox("No queries found for volunteer " & RVLID)
                        Exit Sub
                    End If

                    'RUN REPORT SEPERATING BY VISIT 

                    Dim OK As New ReportViewer
                    OK.Visible = True
                    OK.ReportViewer1.Visible = True
                    OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "Query_Management_Tool.PrintReport.rdlc"
                    OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet", _
                                                               dt))
                    OK.ReportViewer1.RefreshReport()

                Catch ex As Exception

                    MsgBox(ex.Message)
                    Exit Sub

                End Try

            Case "Button104"
                Call Saver(AddQuery.NewQueryGrid3)


        End Select

    End Sub


    Private Function CheckDates() As Boolean

        Dim dater1, dater2 As Date
        dater1 = Form1.DateTimePicker1.Value
        dater2 = Form1.DateTimePicker2.Value

        If dater1 >= dater2 Then
            MsgBox("'Date To' must be greater than 'Date From'")
            CheckDates = False
        Else
            CheckDates = True
        End If

    End Function

    Public Sub AddControls(WhichControl As Control)

        For Each Control In WhichControl.Controls

            If (TypeOf Control Is Button) Then
                Dim But As Button = Control
                AddHandler But.Click, AddressOf ButtonSpecifics
            End If


            If Control.HasChildren Then
                AddControls(Control)
            End If

        Next

    End Sub
End Module

