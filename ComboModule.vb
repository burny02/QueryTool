Module ComboModule

    Public Sub GenericCombo(sender As Object, e As EventArgs)

        If sender.SelectedValue.ToString = vbNullString Then Exit Sub

        If OverClass.UnloadData() = True Then Exit Sub
        OverClass.ResetCollection()
        Call SubCombo(sender)


    End Sub

    Private Sub SubCombo(sender As ComboBox)

        Select Case sender.Name.ToString

            'Case "ComboBox4"
            'StartCombo(Form1.ComboBox3)

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

        End Select

        ComboRefreshData(ctl)

    End Sub

    Public Sub ComboRefreshData(sender As ComboBox)

        Dim Grid As DataGridView = Nothing

        Select Case sender.Name.ToString()


            Case "ComboBox1"
                Grid = Form1.DataGridView2

            Case "ComboBox2"
                Grid = Form1.DataGridView3

            Case "ComboBox4"
                Grid = Form1.DataGridView3

        End Select


        If Not IsNothing(Grid) Then Call Form1.Specifics(Grid)

    End Sub


End Module
