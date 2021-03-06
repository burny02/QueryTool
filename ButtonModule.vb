﻿Imports Microsoft.Reporting.WinForms

Module ButtonModule

    Public Sub ButtonSpecifics(sender As Object, e As EventArgs)

        Dim ctl As Object = Nothing

        Select Case sender.name.ToString

            Case "Button1"
                Dim Trd As New Threading.Thread(Sub() ExportExcel("SELECT * FROM RawResponse", False))
                Trd.Start()

            Case "Button2"
                Dim Trd As New Threading.Thread(Sub() ExportExcel("SELECT * FROM ResponsePerPerson", False))
                Trd.Start()

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
                Dim Trd As New Threading.Thread(Sub() ExportExcel("SELECT * FROM AllQueries", False))
                Trd.Start()

            Case "Button6"
                Dim Trd As New Threading.Thread(Sub() ExportExcel("SELECT * FROM AllQueries WHERE Status='Open'", True))
                Trd.Start()

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

                    Dim TotalCount As Long = Overclass.SELECTCount("SELECT 1 FROM Queries WHERE Status='Open' AND RVLID=" &
                                                                   RVLID)
                    Dim PrintCount As Long = Overclass.SELECTCount("SELECT * FROM PrintOut WHERE Status='Open' AND RVLID=" &
                                                                   RVLID)

                    If TotalCount <> PrintCount Then
                        If MsgBox(TotalCount - PrintCount & " queries were found to be incorrectly coded. Do you want to proceed?",
                                  MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If

                    Dim SqlString As String = "SELECT * FROM PrintOut WHERE Status='Open' AND RVLID=" &
                                                                   RVLID

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

                Form1.NewQueryGrid.EndEdit()
                Dim chkColumns(10) As DataGridViewColumn
                chkColumns(0) = Form1.NewQueryGrid.Columns("RVLID")
                chkColumns(1) = Form1.NewQueryGrid.Columns("Initials")
                chkColumns(2) = Form1.NewQueryGrid.Columns("VisitName")
                chkColumns(3) = Form1.NewQueryGrid.Columns("FormName")
                chkColumns(4) = Form1.NewQueryGrid.Columns("PageNo")
                chkColumns(5) = Form1.NewQueryGrid.Columns("Description")
                chkColumns(6) = Form1.NewQueryGrid.Columns("AssCode")
                chkColumns(7) = Form1.NewQueryGrid.Columns("RespondCode")
                chkColumns(8) = Form1.NewQueryGrid.Columns("Person")
                chkColumns(9) = Form1.NewQueryGrid.Columns("TypeCode")
                chkColumns(10) = Form1.NewQueryGrid.Columns("SiteCode")
                If Overclass.CheckColumns(Form1.NewQueryGrid, chkColumns, chkColumns) = False Then Call Saver(Form1.NewQueryGrid)

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

