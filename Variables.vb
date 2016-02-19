Imports TemplateDB
Module Variables

    Public Overclass As OverClass
    Private Const TablePath As String = "M:\VOLUNTEER SCREENING SERVICES\Systems\Query_Management_Tool\Backend.accdb"
    Private Const PWord As String = "RetroRetro*1"
    Private Const Connect2 As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & TablePath & ";Jet OLEDB:Database Password=" & PWord
    Private Const UserTable As String = "[Users]"
    Private Const UserField As String = "Username"
    Private Const LockTable As String = "[Locker]"
    Private Const AuditTable As String = "[Audit]"
    Private Contact As String = "Craig Tordoff"
    Public Const SolutionName As String = "Query Management Tool"
    Public AccessLevel As Integer = 0
    Public Role As String = vbNullString
    Public AdQry As AddQuery

    Public Sub StartUpCentral()

        Overclass = New OverClass
        Overclass.SetPrivate(UserTable,
                           UserField,
                           LockTable,
                           Contact,
                           Connect2,
                           AuditTable)

        Overclass.LockCheck()

        Overclass.LoginCheck()

        AccessLevel = Overclass.TempDataTable("SELECT Admin FROM [Users] WHERE UserName='" & Overclass.GetUserName & "'").Rows(0).Item(0)
        Role = Overclass.TempDataTable("SELECT Role FROM [Users] WHERE UserName='" & Overclass.GetUserName & "'").Rows(0).Item(0)

        Overclass.AddAllDataItem(Form1)

        For Each ctl In Overclass.DataItemCollection
            If (TypeOf ctl Is Button) Then
                Dim But As Button = ctl
                AddHandler But.Click, AddressOf ButtonSpecifics
            End If
        Next

    End Sub
End Module
