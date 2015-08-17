﻿Imports Microsoft.Reporting.WinForms

Module ButtonModule

    Public Sub ButtonSpecifics(sender As Object, e As EventArgs)

        Dim ctl As Object = Nothing

        Select Case sender.name.ToString

            Case "Button1"
                Call Saver(Form1.DataGridView2)

            Case "Button2"
                Call UploadCSV()

            Case "Button3"
                Call Saver(Form1.DataGridView3)

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
                Call ExportExcel("SELECT a.*, c.* " & _
                        "FROM (((((Queries a INNER JOIN Study b ON a.Study=b.StudyCode) " & _
                        "INNER JOIN QueryCodes c ON a.QueryID=c.QueryID) " & _
                        "INNER JOIN GroupCode d ON b.CodeList=d.ListID) " & _
                        "INNER JOIN TypeCode e ON b.CodeList=e.ListID) " & _
                        "INNER JOIN SiteCode f ON b.CodeList=f.ListID) " & _
                        "WHERE f.code=c.SiteCode AND c.RespondCode=d.code AND TypeCode=e.code" _
                         , Form1.ComboBox3.SelectedValue.ToString, False)

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
                OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "QueryTool.AvgResponse.rdlc"
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
                OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "QueryTool.Totals.rdlc"
                OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet", _
                                                          Overclass.TempDataTable("SELECT * FROM Totals " & _
                                                                                  "WHERE FilterDate Between " & Overclass.SQLDate(Form1.DateTimePicker1.Value) & _
                                                                                  " AND " & Overclass.SQLDate(Form1.DateTimePicker2.Value))))
                OK.ReportViewer1.RefreshReport()

            Case "Button9"
                CheckDates()
                Dim OK As New ReportViewer
                OK.Visible = True
                OK.ReportViewer1.Visible = True
                OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "QueryTool.QCTeam.rdlc"
                OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet", _
                                                          Overclass.TempDataTable("SELECT * FROM QCTeam " & _
                                                                                  "WHERE FilterDate Between " & Overclass.SQLDate(Form1.DateTimePicker1.Value) & _
                                                                                  " AND " & Overclass.SQLDate(Form1.DateTimePicker2.Value))))
                OK.ReportViewer1.RefreshReport()

            Case "Button10"
                CheckDates()
                Dim OK As New ReportViewer
                OK.Visible = True
                OK.ReportViewer1.Visible = True
                OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "QueryTool.Types.rdlc"
                OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet", _
                                                          Overclass.TempDataTable("SELECT * FROM Types " & _
                                                                                  "WHERE FilterDate Between " & Overclass.SQLDate(Form1.DateTimePicker1.Value) & _
                                                                                  " AND " & Overclass.SQLDate(Form1.DateTimePicker2.Value))))
                OK.ReportViewer1.RefreshReport()

            Case "Button11"
                CheckDates()
                Dim OK As New ReportViewer
                OK.Visible = True
                OK.ReportViewer1.Visible = True
                OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "QueryTool.Responders.rdlc"
                OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet", _
                                                          Overclass.TempDataTable("SELECT * FROM Responders " & _
                                                                                  "WHERE FilterDate Between " & Overclass.SQLDate(Form1.DateTimePicker1.Value) & _
                                                                                  " AND " & Overclass.SQLDate(Form1.DateTimePicker2.Value))))
                OK.ReportViewer1.RefreshReport()

            Case "Button12"
                CheckDates()
                Dim OK As New ReportViewer
                OK.Visible = True
                OK.ReportViewer1.Visible = True
                OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "QueryTool.QCIndividual.rdlc"
                OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet", _
                                                          Overclass.TempDataTable("SELECT * FROM QCIndividual " & _
                                                                                  "WHERE FilterDate Between " & Overclass.SQLDate(Form1.DateTimePicker1.Value) & _
                                                                                  " AND " & Overclass.SQLDate(Form1.DateTimePicker2.Value))))
                OK.ReportViewer1.RefreshReport()

            Case "Button13"
                CheckDates()
                Dim OK As New ReportViewer
                OK.Visible = True
                OK.ReportViewer1.Visible = True
                OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "QueryTool.ToolUsage.rdlc"
                OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet", _
                                                          Overclass.TempDataTable("SELECT * FROM ToolUsage " & _
                                                                                  "WHERE FilterDate Between " & Overclass.SQLDate(Form1.DateTimePicker1.Value) & _
                                                                                  " AND " & Overclass.SQLDate(Form1.DateTimePicker2.Value))))
                OK.ReportViewer1.RefreshReport()

            Case "Button14"
                CheckDates()
                Dim OK As New ReportViewer
                OK.Visible = True
                OK.ReportViewer1.Visible = True
                OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "QueryTool.Deviations.rdlc"
                OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet", _
                                                          Overclass.TempDataTable("SELECT * FROM Deviations " & _
                                                                                  "WHERE FilterDate Between " & Overclass.SQLDate(Form1.DateTimePicker1.Value) & _
                                                                                  " AND " & Overclass.SQLDate(Form1.DateTimePicker2.Value))))
                OK.ReportViewer1.RefreshReport()

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

End Module
