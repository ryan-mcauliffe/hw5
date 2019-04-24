
Imports System.Data.OleDb
Public Class Database
        Dim rdbSQL As String
        Private rdbDataAdapter As New OleDbDataAdapter
        Private rdbDataConnection As New OleDbConnection
        Private rdbConnectionString As String
        Private rdbDataSet As New DataSet
        Private rdbCommand As New OleDbCommand
        Private rdbTableName As String
        Public Sub New()

        End Sub
        Public Sub RunSql(rdbSQL, rdbConnectionString, rdbDataSet, rdbTableName)
            'RDB: First we build the Database Connection Object
            rdbDataConnection.ConnectionString = rdbConnectionString

            'RDB: Build a Database Command
            rdbCommand = rdbDataConnection.CreateCommand()
            rdbCommand.CommandText = rdbSQL

            ' RDB: Build the Data Adapter
            rdbDataAdapter.SelectCommand = rdbCommand
            rdbDataAdapter.Fill(rdbDataSet, rdbTableName)

        End Sub
    End Class


