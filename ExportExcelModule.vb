Imports System.IO
Module ExportExcelModule
    Public Sub ExportExcel(SQLCode As String, Send As Boolean)

        If Send = False Then
            Dim dt As New DataTable
            Dim da As OleDb.OleDbDataAdapter = Overclass.NewDataAdapter(SQLCode)
            da.Fill(dt)

            da = Nothing
            Dim i As Integer
            Dim j As Integer

            Dim xlApp As Object
            xlApp = CreateObject("Excel.Application")
            With xlApp
                .Visible = False
                .Workbooks.Add()
                .Sheets("Sheet1").Select()

                Dim Sheet = xlApp.activesheet

                'Add column heading
                For i = 1 To dt.Columns.Count
                    Sheet.Cells(1, i).Value = dt.Columns(i - 1).ColumnName
                Next i

                'Add Rows
                For i = 0 To dt.Rows.Count - 1
                    For j = 0 To dt.Columns.Count - 1
                        Sheet.Cells(i + 2, j + 1) = dt.Rows(i).Item(j)
                    Next j
                Next i

                .Cells.EntireColumn.AutoFit()
                Sheet.Range("$A$1:  $Z$1").AutoFilter()

            End With

            xlApp.Visible = True

            dt = Nothing
            da = Nothing



        End If

        If Send = True Then

            Dim OutApp
            Dim objOutlookMsg
            Dim Inspector

            Try

                OutApp = CreateObject("Outlook.Application")
                objOutlookMsg = OutApp.CreateItem(0)
                Inspector = objOutlookMsg.GetInspector

                Dim CountTable As DataTable = Overclass.TempDataTable("SELECT * FROM ExportExcelCount ORDER BY Study, Site")

                Dim TableView As String = vbNullString

                TableView = "<head>
                         <title>HTML Table Cellpadding</title>
                        </head>
                        <body>
                        <table border = ""1"" cellpadding=""5"" cellspacing=""5"">
                        <tr>
                        <th>Study</th>
                        <th>Site</th>
                        <th>Total Queries</th>
                        <th>QC Total</th>
                        <th>DM Total</th>
                        <th>Overdue Queries</th>
                        <th>Priority 1</th>
                        <th>Priority 2</th>
                        </tr>"

                For Each row As DataRow In CountTable.Rows
                    TableView = TableView &
                                "<tr>
                            <td>" & row.Item("Study") & "</td>
                            <td>" & row.Item("Site") & "</td>
                            <td>" & row.Item("Tot_No") & "</td>
                            <td>" & row.Item("QC_Total") & "</td>
                            <td>" & row.Item("DM_Total") & "</td>
                            <td>" & row.Item("Overdue") & "</td>
                            <td>" & row.Item("PriorityOne") & "</td>
                            <td>" & row.Item("PriorityTwo") & "</td>
                            </tr>"
                Next

                TableView = TableView & "</table>
                                    </body>"


                Dim Link As String = "<a href='M:\VOLUNTEER SCREENING SERVICES\Systems\Query_Management_Tool\Query Management Tool.application'>" &
                                        "Click Here</a>"

                objOutlookMsg.Subject = "Open Queries"

                objOutlookMsg.HTMLBody = "Dear All" & "<br/>" & "<br/>" &
                                            "The following queries are currently open:  " & "<br/>" & "<br/>" &
                                            TableView & "<br/>" &
                                            "Link to Query Tool: " & Link & "<br/>" & "<br/>" &
                                            "Many thanks"


                objOutlookMsg.Display()
                Inspector.activate()



            Catch ex As Exception
                objOutlookMsg = Nothing
                MsgBox(ex.Message)
            Finally
                OutApp = Nothing
                Inspector = Nothing
            End Try

        End If

    End Sub


End Module
