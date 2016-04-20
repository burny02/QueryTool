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
            If MsgBox("Would you Like to attach the spreadsheet to an email?", vbYesNo) = vbNo Then Send = False
        End If


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
        Dim r As Long
        numrow = dt.Rows.Count + 1
        dt = Nothing
        da = Nothing

        If Send = False Then xlApp.Visible = True

        If Send = True Then

            For r = 2 To numrow
                If xlApp.activesheet.Range("$A$" & r).Value < Date.Now Then
                    xlApp.activesheet.Range("$A$" & r & ":$Z$" & r).Font.ColorIndex = 3
                End If

            Next r

            Dim Namer As String

            Namer = Study & " Queries " & Format(Now(), "dd-mmm-yyyy")

            If Not Directory.Exists("C:\DBS") Then MkDir("C:\DBS")

            xlApp.DisplayAlerts = False

            'SAVES FILE USING THE VARIABLE BOOKNAME AS FILENAME
            xlApp.ActiveWorkbook.SaveAs("C:\DBS\" & Namer & ".xlsx")

            xlApp.DisplayAlerts = True

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



            objOutlookMsg.Subject = Study & " Queries"

            objOutlookMsg.HTMLBody = "Dear All" & "<br/>" & "<br/>" &
                                        "Please find attached the latest query spreadsheet:  " & "<br/>" & "<br/>" &
                                        TableString & "<br/>" &
                                        "Please filter by columns allocation, site, group and priority as required." & "<br/>" & "<br/>" &
                                        "Overdue queries are marked in red." & "<br/>" & "<br/>" &
                                        "Many thanks"


            objOutlookMsg.Display()
            objOutlookMsg.Attachments.Add(xlApp.ActiveWorkbook.fullname.ToString)

            OutApp = Nothing
            xlApp.Quit()

        End If

    End Sub

End Module
