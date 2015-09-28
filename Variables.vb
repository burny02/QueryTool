Imports TemplateDB
Module Variables

    Public Overclass As OverClass
    Private Const TablePath As String = "M:\VOLUNTEER SCREENING SERVICES\DavidBurnside\Queries\Backend3.accdb"
    Private Const PWord As String = "Crypto*Dave02"
    Private Const Connect2 As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & TablePath & ";Jet OLEDB:Database Password=" & PWord
    Private Const UserTable As String = "[Users]"
    Private Const UserField As String = "Username"
    Private Const LockTable As String = "[Locker]"
    Private Const ActiveUserTable As String = "[ActiveUsers]"
    Private Contact As String = "Craig Tordoff"
    Public Const SolutionName As String = "Query Tool"
    Public AccessLevel As Integer = 0
    Public Role As String = vbNullString

    Public Sub StartUpCentral()

        overclass = New OverClass
        overclass.SetPrivate(UserTable, _
                           UserField, _
                           LockTable, _
                           Contact, _
                           Connect2, _
                           ActiveUserTable)

        OverClass.LockCheck()

        Overclass.LoginCheck()

        AccessLevel = Overclass.TempDataTable("SELECT Admin FROM [Users] WHERE UserName='" & Overclass.GetUserName & "'").Rows(0).Item(0)
        Role = Overclass.TempDataTable("SELECT Role FROM [Users] WHERE UserName='" & Overclass.GetUserName & "'").Rows(0).Item(0)

        OverClass.AddAllDataItem(Form1)

        For Each ctl In OverClass.DataItemCollection
            If (TypeOf ctl Is ComboBox) Then
                Dim Com As ComboBox = ctl
                AddHandler Com.SelectionChangeCommitted, AddressOf GenericCombo
            End If
        Next
        For Each ctl In OverClass.DataItemCollection
            If (TypeOf ctl Is Button) Then
                Dim But As Button = ctl
                AddHandler But.Click, AddressOf ButtonSpecifics
            End If
        Next

    End Sub
End Module
