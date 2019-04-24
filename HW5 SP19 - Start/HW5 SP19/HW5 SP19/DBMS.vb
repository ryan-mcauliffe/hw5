
Imports System.Data.OleDb                             'RDB: We use the OleDB namespace for an MS Access namespace
'**************************************************************************************************************************
Public Class DBMS
    'ADO.NET Data Provider Objects
    Private rdbDataAdapter As New OleDbDataAdapter    'RDB: ADO.NET DataAdapter Object -  The DataAdapter facilitates the conversion from the external Database to the internal DataSet
    Private rdbConnection As New OleDbConnection      'RDB: ADO.NET Connection Object property -  
    Private rdbConnectionString As String             'RDB: ADO.NET Connection Object property -  The connection string is found in the app.config file.  This string points to the external DataBase
    Private rdbSQL As String                          'RDB: ADO.NET Command Object property - This string is declared for any SQL statements that we will write
    Private rdbDataSet As New DataSet                 'RDB: ADO.NET DataSet Object - The external Database will be transformed into an internal DataSetDim orderList As New List(Of SalesOrder)
    Private rdbCommand As New OleDbCommand            'RDB: ADO.NET DataSet Command Object - sql commands are conveyed via a Command object
    Private rdbTableName As String                    'RDB: sql Commands are executed on tables
    '**************************************************************************************************************************
    Public Sub New()
        ' This sub creates a new database object

    End Sub
    '**************************************************************************************************************************
    Public Sub RunSQL(ByVal rdbTableName, ByVal rdbSQL, ByVal rdbDataSet, ByVal rdbConnectionString)
        '
        'RDB:  First we build the Database Connection Object
        rdbConnection.ConnectionString = rdbConnectionString      'RDB:  We use the Connection String that is found in the app.config file to connect to the DataBase
        rdbCommand = rdbConnection.CreateCommand()
        '
        'RDB: Now we build the Command that will be sent to the Data Adapter
        rdbCommand.CommandText = rdbSQL                            'RDB: The sql is embedded in an object called a Command
        '
        'RDB: Now we build the Data Adapter
        rdbDataAdapter.SelectCommand = rdbCommand                  'RDB: The DataAdaptor executes the query Command
        'RDB: The Fill method of the DataAdaptor fills tables in the DataSet from data from the DataBase
        'RDB: Note up to this point in the code, the DataSet was declared, but it had no tables, rows and columns defined
        'RDB: The Fill method executes the SQL that defines the table and columns and fills the rows in the DataSet 
        rdbDataAdapter.Fill(rdbDataSet, rdbTableName)              'RDB: The Fill method of the DataAdaptor fills tables in the DataSet from data from the DataBase

    End Sub
End Class
