Imports System.IO
Module ExportExcelModule
    Public Sub ExportExcel(SQLCode As String, Study As String, Send As Boolean, Optional Check As Boolean = True)

        Dim NumberWrong As Long
        If Check = True Then
            NumberWrong = Overclass.SELECTCount("Select Study, DisplayName, QueryID, SiteCode, TypeCode, RespondCode, RVLID, " &
                              "VisitName, FormName, Description, Status, Person FROM IncorrectQueries " &
                              "Where Study='" & Study & "'")

            If NumberWrong <> 0 Then
                If MsgBox(NumberWrong & " bad/empty codes were found and will be missing from report. Do you wish to proceed?", vbYesNo) = vbNo Then Exit Sub
            End If

        End If

        If Send = True Then
            If MsgBox("Would you to email queries?", vbYesNo) = vbNo Then Send = False
        End If

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

                'Add column heading
                For i = 1 To dt.Columns.Count
                    xlApp.activesheet.Cells(1, i).Value = dt.Columns(i - 1).ColumnName
                Next i

                'Add Rows
                For i = 0 To dt.Rows.Count - 1
                    For j = 0 To dt.Columns.Count - 1
                        xlApp.activesheet.Cells(i + 2, j + 1) = dt.Rows(i).Item(j)
                    Next j
                Next i

                xlApp.Cells.EntireColumn.AutoFit()
                .activesheet.Range("$A$1: $Z$1").AutoFilter()

            End With

            Dim numrow As Long
            numrow = dt.Rows.Count + 1
            dt = Nothing
            da = Nothing

            xlApp.Visible = True

        End If

        If Send = True Then

            Dim OutApp = CreateObject("Outlook.Application")
            Dim objOutlookMsg = OutApp.CreateItem(0)

            OutApp = CreateObject("Outlook.Application")
            objOutlookMsg = OutApp.CreateItem(0)

            Dim CountTable As DataTable = Overclass.TempDataTable("SELECT Site, Priority, Queriess, Overdue FROM ExportExcelCount WHERE a.Study='" & Study & "'" &
                                                                  " ORDER BY Site, Priority")

            Dim TableString As String = vbNullString

            For Each row As DataRow In CountTable.Rows
                TableString = TableString & row.Item("Site")
                TableString = TableString & " - Priority " & row.Item("Priority")
                TableString = TableString & " (" & row.Item("Queriess") & " - " & row.Item("Overdue") & ")"
                TableString = TableString & "<br/>"
            Next


            Dim Link As String = "<a href='M:\VOLUNTEER SCREENING SERVICES\Systems\Query_Management_Tool\Query Management Tool.application'>" &
                                    "Ctrl + Click Here</a>"

            objOutlookMsg.Subject = Study & " Queries"

            objOutlookMsg.HTMLBody = "Dear All" & "<br/>" & "<br/>" &
                                        "The following queries are currently open:  " & "<br/>" & "<br/>" &
                                        TableString & "<br/>" &
                                        "Link to Query Tool: " & Link & "<br/>" & "<br/>" &
                                        "Many thanks"


            objOutlookMsg.Display()

            OutApp = Nothing

        End If

    End Sub

End Module
