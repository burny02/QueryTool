Imports Microsoft.Reporting.WinForms

Module ButtonModule

    Public Sub ButtonSpecifics(sender As Object, e As EventArgs)

        Dim ctl As Object = Nothing

        Select Case sender.name.ToString

            Case "Button1"
                Call Saver(Form1.DataGridView2)

            Case "Button2"
                Call UploadCSV()

            Case "Button3"
                CheckDates()
                Dim OK As New ReportViewer
                OK.Visible = True
                OK.ReportViewer1.Visible = True
                OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "Query_Management_Tool.QueryPerDay.rdlc"
                OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("Dave0",
                                                          Overclass.TempDataTable("Select * FROM QueriesPerDay " &
                                                                                  "WHERE FilterDate Between " & Overclass.SQLDate(Form1.DateTimePicker1.Value) &
                                                                                  " And " & Overclass.SQLDate(Form1.DateTimePicker2.Value))))
                OK.ReportViewer1.RefreshReport()

            Case "Button5"
                Call ExportExcel("SELECT * FROM AllQueries", Nothing, False, False)

            Case "Button6"
                Call ExportExcel("SELECT dateadd('d',QueryAgeLimit,CreateDate) AS DueDate," &
                         "Person as [Allocated To], Site, Group, RVLID, " &
                        "FormName, Description, Priority " &
                        "FROM (((((Queries a INNER JOIN Study b ON a.Study=b.StudyCode) " &
                        "INNER JOIN QueryCodes c ON a.QueryID=c.QueryID) " &
                        "INNER JOIN GroupCode d ON b.CodeList=d.ListID) " &
                        "INNER JOIN TypeCode e ON b.CodeList=e.ListID) " &
                        "INNER JOIN SiteCode f ON b.CodeList=f.ListID) " &
                        "WHERE Status='Open' AND Study='" & Form1.FilterCombo6.SelectedValue.ToString & "' " &
                        " AND f.code=c.SiteCode AND c.RespondCode=d.code AND TypeCode=e.code" _
                         , Form1.FilterCombo6.SelectedValue.ToString, True)

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
                AdQry.TabControl1.Controls.Remove(AdQry.TabPage2)
                AdQry.ShowDialog()




            Case "Button17"

                'SELECT RVLID FROM CURRENTDATA SET
                Dim RVLID As Long

                Try
                    RVLID = InputBox("Please input a volunteers subject ID ", "Subject ID")
                Catch ex As Exception
                    MsgBox("Error - Expected number")
                    Exit Sub
                End Try


                Try

                    Dim TotalCount As Long = Overclass.SELECTCount("SELECT 1 FROM Queries WHERE Status='Open' AND RVLID='" &
                                                                   RVLID & "'")
                    Dim PrintCount As Long = Overclass.SELECTCount("SELECT * FROM PrintOut WHERE Status='Open' AND RVLID='" &
                                                                   RVLID & "'")

                    If TotalCount <> PrintCount Then
                        If MsgBox(TotalCount - PrintCount & " queries were found to be incorrectly coded. Do you want to proceed?",
                                  MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If

                    Dim SqlString As String = "SELECT * FROM PrintOut WHERE Status='Open' AND RVLID='" &
                                                                   RVLID & "'"

                    Dim dt As DataTable = Overclass.TempDataTable(SqlString)

                    If dt.Rows.Count = 0 Then
                        MsgBox("No open coded queries found for volunteer")
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

            Case "Button202"
                Call Saver(AddQuery.NewQueryGrid2)

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

