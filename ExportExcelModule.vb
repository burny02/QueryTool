Imports System.IO
Module ExportExcelModule
    Public Sub ExportExcel(SQLCode As String, Study As String, Send As Boolean)

        Dim NumberWrong As Long = Overclass.SELECTCount("Select Study, DisplayName, QueryID, SiteCode, TypeCode, RespondCode, RVLID, " &
                              "VisitName, FormName, Description, Status, Person FROM IncorrectQueries " &
                              "Where Study='" & Study & "'")

        If NumberWrong <> 0 Then
            If MsgBox(NumberWrong & " bad/empty codes were found And will be missing from report. Do you wish to proceed?", vbYesNo) = vbNo Then Exit Sub
        End If

        Dim WantSend As Boolean = False

        If Send = True Then
            If MsgBox("Would you Like to attach the spreadsheet to an email?", vbYesNo) = vbYes Then WantSend = True
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

        If WantSend = False Then xlApp.Visible = True

        If WantSend = True Then

            For r = 2 To numrow
                If xlApp.activesheet.Range("$A$" & r).Value < Date.Now Then
                    xlApp.activesheet.Range("$A$" & r & ":$Z$" & r).Font.ColorIndex = 3
                Else
                    Exit For
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

            Dim MSite As Long
            Dim WSite As Long
            Dim QSite As Long
            Dim OvMSite As Long
            Dim OvWSite As Long
            Dim OvQSite As Long

            OutApp = CreateObject("Outlook.Application")
            objOutlookMsg = OutApp.CreateItem(0)

            MSite = Overclass.SELECTCount("SELECT a.QueryID FROM (((QueryCodes a INNER JOIN Queries b on a.QueryID=b.QueryID) " & _
                                "INNER JOIN Study c on b.Study=c.StudyCode) " & _
                                "INNER JOIN SiteCode d on c.Codelist=d.ListID) " & _
                                "WHERE Status='Open' AND Site='MAN' " & _
                                "AND SiteCode=Code AND Study='" & Study & "'")

            WSite = Overclass.SELECTCount("SELECT a.QueryID FROM (((QueryCodes a INNER JOIN Queries b on a.QueryID=b.QueryID) " & _
                                "INNER JOIN Study c on b.Study=c.StudyCode) " & _
                                "INNER JOIN SiteCode d on c.Codelist=d.ListID) " & _
                                "WHERE Status='Open' AND Site='WHC' " & _
                                "AND SiteCode=Code AND Study='" & Study & "'")

            QSite = Overclass.SELECTCount("SELECT a.QueryID FROM (((QueryCodes a INNER JOIN Queries b on a.QueryID=b.QueryID) " & _
                                "INNER JOIN Study c on b.Study=c.StudyCode) " & _
                                "INNER JOIN SiteCode d on c.Codelist=d.ListID) " & _
                                "WHERE Status='Open' AND Site='Quarantine' " & _
                                "AND SiteCode=Code AND Study='" & Study & "'")

            OvMSite = Overclass.SELECTCount("SELECT a.QueryID FROM (((QueryCodes a INNER JOIN Queries b on a.QueryID=b.QueryID) " & _
                                "INNER JOIN Study c on b.Study=c.StudyCode) " & _
                                "INNER JOIN SiteCode d on c.Codelist=d.ListID) " & _
                                "WHERE Status='Open' AND Site='MAN' " & _
                                "AND dateadd('d',QueryAgeLimit,CreateDate)<Date()  " & _
                                "AND SiteCode=Code AND Study='" & Study & "'")

            OvWSite = Overclass.SELECTCount("SELECT a.QueryID FROM (((QueryCodes a INNER JOIN Queries b on a.QueryID=b.QueryID) " & _
                                "INNER JOIN Study c on b.Study=c.StudyCode) " & _
                                "INNER JOIN SiteCode d on c.Codelist=d.ListID) " & _
                                "WHERE Status='Open' AND Site='WHC' " & _
                                "AND dateadd('d',QueryAgeLimit,CreateDate)<Date()  " & _
                                "AND SiteCode=Code AND Study='" & Study & "'")

            OvQSite = Overclass.SELECTCount("SELECT a.QueryID FROM (((QueryCodes a INNER JOIN Queries b on a.QueryID=b.QueryID) " & _
                                "INNER JOIN Study c on b.Study=c.StudyCode) " & _
                                "INNER JOIN SiteCode d on c.Codelist=d.ListID) " & _
                                "WHERE Status='Open' AND Site='Quarantine' " & _
                                "AND dateadd('d',QueryAgeLimit,CreateDate)<Date()  " & _
                                "AND SiteCode=Code AND Study='" & Study & "'")



            objOutlookMsg.Subject = Study & " Queries"

            objOutlookMsg.HTMLBody = "Dear All" & "<br/>" & "<br/>" & _
                                        "Please find attached the latest query spreadsheet:" & "<br/>" & "<br/>" & "<br/>" & _
                                        "MAN - " & MSite & "(" & OvMSite & " overdue)" & "<br/>" & _
                                        "WHC - " & WSite & "(" & OvWSite & " overdue)" & "<br/>" & _
                                        "VCU - " & QSite & "(" & OvQSite & " overdue)" & "<br/>" & "<br/>" & "<br/>" & _
                                        "Please filter by columns B,C,D as required." & "<br/>" & "<br/>" & _
                                        "Overdue queries are marked in red." & "<br/>" & "<br/>" & _
                                        "Many thanks"


            objOutlookMsg.Display()
            objOutlookMsg.Attachments.Add(xlApp.ActiveWorkbook.fullname.ToString)

            OutApp = Nothing
            xlApp.Quit()

        End If

    End Sub

End Module
