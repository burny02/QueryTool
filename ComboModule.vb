Module ComboModule

    Public Sub GenericCombo(sender As Object, e As EventArgs)

        If OverClass.UnloadData() = True Then Exit Sub
        OverClass.ResetCollection()
        Call SubCombo(sender)


    End Sub

    Private Sub SubCombo(sender As ComboBox)

        Select Case sender.Name.ToString

            Case "ComboBox1", "ComboBox8", "ComboBox7", "ComboBox6", "ComboBox5", "ComboBox9"

                Form1.Specifics(Form1.DataGridView2)
                StartCombo(Form1.ComboBox8)
                StartCombo(Form1.ComboBox7)
                StartCombo(Form1.ComboBox6)
                StartCombo(Form1.ComboBox5)
                StartCombo(Form1.ComboBox9)

            Case "ComboBox101", "ComboBox102", "ComboBox103", "ComboBox105"
                Form1.Specifics(AdQry.NewQueryGrid)
                StartCombo(AdQry.ComboBox102)
                StartCombo(AdQry.ComboBox103)
                StartCombo(AdQry.ComboBox105)


            Case Else
                ComboRefreshData(sender)

        End Select

    End Sub

    Public Sub StartCombo(ctl As ComboBox)

        Select Case ctl.Name.ToString()


            Case "ComboBox1"
                ctl.DataSource = overclass.TempDataTable("SELECT DisplayName, StudyCode FROM Study ORDER BY StudyCode ASC")
                ctl.ValueMember = "StudyCode"
                ctl.DisplayMember = "DisplayName"

            Case "ComboBox2"
                ctl.DataSource = Overclass.TempDataTable("SELECT DisplayName, StudyCode FROM Study ORDER BY StudyCode ASC")
                ctl.ValueMember = "StudyCode"
                ctl.DisplayMember = "DisplayName"

            Case "ComboBox3"
                ctl.DataSource = overclass.TempDataTable("SELECT DisplayName, StudyCode FROM Study ORDER BY StudyCode ASC")
                ctl.ValueMember = "StudyCode"
                ctl.DisplayMember = "DisplayName"

            Case "ComboBox4"
                ctl.DataSource = Overclass.TempDataTable("SELECT Site FROM Status GROUP BY Site ORDER BY Site ASC")
                ctl.ValueMember = "Site"
                ctl.DisplayMember = "Site"

            Case "ComboBox8"

                If ctl.SelectedValue <> "" Then Exit Sub

                Dim dt As DataTable = Overclass.TempDataTable("SELECT DISTINCT * FROM ( " & _
                                            "SELECT '' AS RVLID FROM Queries " & _
                                            "UNION ALL " & _
                                            "SELECT RVLID " & _
                                            "FROM Queries WHERE Study='" & Form1.ComboBox1.SelectedValue.ToString & "'" & _
                                            " AND ISNull(RVLID)=False) " & _
                                            "ORDER BY RVLID ASC")



                ctl.DataSource = dt
                ctl.DisplayMember = "RVLID"
                ctl.ValueMember = "RVLID"

            Case "ComboBox7"

                If ctl.SelectedValue <> "" Then Exit Sub

                Dim dt As DataTable = Overclass.TempDataTable("SELECT DISTINCT * FROM ( " & _
                                                              "SELECT '' AS Initials FROM Queries " & _
                                                              "UNION ALL " & _
                                                              "SELECT Initials " & _
                                            "FROM Queries WHERE Study='" & Form1.ComboBox1.SelectedValue.ToString & "'" & _
                                            " AND ISNull(Initials)=False) " & _
                                            "ORDER BY Initials ASC")



                ctl.DataSource = dt
                ctl.DisplayMember = "Initials"
                ctl.ValueMember = "Initials"

            Case "ComboBox6"

                If ctl.SelectedValue <> "" Then Exit Sub

                Dim dt As DataTable = Overclass.TempDataTable("SELECT DISTINCT * FROM ( " & _
                                                              "SELECT '' AS Status FROM Queries " & _
                                                              "UNION ALL " & _
                                                              "SELECT Status " & _
                                            "FROM Queries WHERE Study='" & Form1.ComboBox1.SelectedValue.ToString & "'" & _
                                            " AND ISNull(Status)=False) " & _
                                            "ORDER BY Status ASC")



                ctl.DataSource = dt
                ctl.DisplayMember = "Status"
                ctl.ValueMember = "Status"

            Case "ComboBox5"

                If ctl.SelectedValue <> "" Then Exit Sub

                Dim dt As DataTable = Overclass.TempDataTable("SELECT DISTINCT * FROM ( " & _
                                                              "SELECT '' AS VisitName FROM Queries " & _
                                                              "UNION ALL " & _
                                                              "SELECT VisitName " & _
                                            "FROM Queries WHERE Study='" & Form1.ComboBox1.SelectedValue.ToString & "'" & _
                                            " AND ISNull(VisitName)=False) " & _
                                            "ORDER BY VisitName ASC")



                ctl.DataSource = dt
                ctl.DisplayMember = "VisitName"
                ctl.ValueMember = "VisitName"

            Case "ComboBox9"

                If ctl.SelectedValue <> "" Then Exit Sub

                Dim dt As DataTable = Overclass.TempDataTable("SELECT DISTINCT * FROM ( " & _
                                                              "SELECT '' AS SiteCode FROM Queries " & _
                                                              "UNION ALL " & _
                                                              "SELECT SiteCode " & _
                                            "FROM QueryCodes) ORDER BY SiteCode ASC")



                ctl.DataSource = dt
                ctl.DisplayMember = "SiteCode"
                ctl.ValueMember = "SiteCode"

            Case "ComboBox101", "ComboBox107", "ComboBox108"

                ctl.DataSource = Overclass.TempDataTable("SELECT DisplayName, StudyCode " & _
                                                         "FROM Study ORDER BY StudyCode ASC")
                ctl.DisplayMember = "DisplayName"
                ctl.ValueMember = "StudyCode"

            Case "ComboBox102"

                If ctl.SelectedValue <> "" Then Exit Sub

                Dim dt As DataTable = Overclass.TempDataTable("SELECT DISTINCT * FROM ( " & _
                                            "SELECT '' AS RVLID FROM Queries " & _
                                            "UNION ALL " & _
                                            "SELECT RVLID " & _
                                            "FROM Queries " & _
                                            "WHERE IsNull(RVLID)=False AND RVLID IN ( SELECT RVLID FROM ( " _
                                            & Overclass.CurrentDataAdapter.SelectCommand.CommandText & _
                                            "))) ORDER BY RVLID  ASC")



                ctl.DataSource = dt
                ctl.DisplayMember = "RVLID"
                ctl.ValueMember = "RVLID"

            Case "ComboBox103"

                If ctl.SelectedValue <> "" Then Exit Sub

                Dim dt As DataTable = Overclass.TempDataTable("SELECT DISTINCT * FROM ( " & _
                                                              "SELECT '' AS Initials FROM Queries " & _
                                                              "UNION ALL " & _
                                                              "SELECT Initials " & _
                                            "FROM Queries " & _
                                            "WHERE IsNull(Initials)=False AND Initials IN ( SELECT Initials FROM ( " _
                                            & Overclass.CurrentDataAdapter.SelectCommand.CommandText & _
                                            "))) ORDER BY Initials  ASC")



                ctl.DataSource = dt
                ctl.DisplayMember = "Initials"
                ctl.ValueMember = "Initials"

            Case "ComboBox105"

                If ctl.SelectedValue <> "" Then Exit Sub

                Dim dt As DataTable = Overclass.TempDataTable("SELECT DISTINCT * FROM ( " & _
                                                              "SELECT '' AS VisitName FROM Queries " & _
                                                              "UNION ALL " & _
                                                              "SELECT VisitName " & _
                                            "FROM Queries " & _
                                            "WHERE IsNull(VisitName)=False AND VisitName IN ( SELECT VisitName FROM ( " _
                                            & Overclass.CurrentDataAdapter.SelectCommand.CommandText & _
                                            "))) ORDER BY VisitName  ASC")



                ctl.DataSource = dt
                ctl.DisplayMember = "VisitName"
                ctl.ValueMember = "VisitName"

        End Select

        ComboRefreshData(ctl)

    End Sub

    Public Sub ComboRefreshData(sender As ComboBox)

        Dim Grid As DataGridView = Nothing

        Select Case sender.Name.ToString()


            Case "ComboBox107"
                Grid = AdQry.NewQueryGrid2

            Case "ComboBox108"
                Grid = AdQry.NewQueryGrid3


        End Select


        If Not IsNothing(Grid) Then Call Form1.Specifics(Grid)

    End Sub


End Module
