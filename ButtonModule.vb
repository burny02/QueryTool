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

        End Select

    End Sub

End Module
