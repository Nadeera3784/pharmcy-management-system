Imports System.Data.SqlClient
Module ModCommonClasses
    Public con As SqlConnection = Nothing
    Public cmd, cmd1 As SqlCommand
    Public rdr As SqlDataReader = Nothing
    Public ds As DataSet
    Public adp As SqlDataAdapter
    Public dtable As DataTable
    Public TempFileNames2 As String
End Module
