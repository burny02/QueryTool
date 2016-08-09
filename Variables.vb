Imports TemplateDB
Module Variables

    Public Overclass As OverClass
    Private Const TablePath As String = "M:\VOLUNTEER SCREENING SERVICES\Systems\Query_Management_Tool\Test.accdb"
    Private Const PWord As String = "RetroRetro*1"
    Private Const Connect2 As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & TablePath & ";Jet OLEDB:Database Password=" & PWord
    Private Const UserTable As String = "[Users]"
    Private Const UserField As String = "Username"
    Private Const AuditTable As String = "[Audit]"
    Private Contact As String = "Craig Tordoff"
    Public Const SolutionName As String = "Query Management Tool"
    Public AccessLevel As Integer = 0
    Public Role As String = vbNullString
    Public RespView As ResponseView
    Public RespondCommands As New List(Of OleDb.OleDbCommand)
    Public CurrentRecord As Long

    Public Sub StartUpCentral()

        Overclass = New OverClass
        Overclass.SetPrivate(UserTable,
                           UserField,
                           Contact,
                           Connect2,
                           AuditTable)

        Dim SQLString(0) As String
        SQLString(0) = "SELECT Admin, Role FROM [Users] WHERE UserName='" & Overclass.GetUserName & "'"

        Dim dt() As DataTable = Overclass.LoginCheck(SQLString)

        AccessLevel = dt(1).Rows(0).Item(0)
        Role = dt(1).Rows(0).Item(1)

        If AccessLevel <> 0 Then
            Overclass.AddAllDataItem(Form1)

            For Each ctl In Overclass.DataItemCollection
                If (TypeOf ctl Is Button) Then
                    Dim But As Button = ctl
                    AddHandler But.Click, AddressOf ButtonSpecifics
                End If
            Next
        End If

    End Sub
End Module
