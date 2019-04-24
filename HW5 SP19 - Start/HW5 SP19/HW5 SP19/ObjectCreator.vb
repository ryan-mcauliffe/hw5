Public Class ObjectCreator
    'RDB: This Class is designed to extract data from database tables and create objects with properties that correspond to the tables
    '    Dim RDBDataSet As New DataSet

    '******************************************************************************
    Public Sub CreateObjectsAndLists()
        'RDB: Initialize the connection string to the database for this project
        Dim HW5ConnectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\HW5 SP19.mdb"
        '    Dim RDBDataSet As New DataSet

        'RDB: We define key variables for the objects that are needed to make a databse connection
        Dim tableName As String
        Dim HW5SQL As String                'RDB: This string will hold the SQL statements that we send to the RunSQL method of the Data Object
        Dim HW5Data As New DBMS             'RDB: We create an object of the Data Class for running SQL that extracts data from the Data base 
        Dim HW5DataSet As New DataSet       'RDB: We create a DataSet object to be able to refer to the DataSet of this project
        '
        'RDB: We set properties of the PeriodData objects from data in the PeriodsData table
        tableName = "PeriodData"
        HW5SQL = "SELECT * FROM " & tableName                               'RDB: This SQL selects all columns from the Meal table in MS Access
        HW5Data.RunSQL(tableName, HW5SQL, HW5DataSet, HW5ConnectionString)  'RDB: We execute the query
        '
        'RDB: We create the objects of the Period Class from the data in the Period table
        For rowIndex As Integer = 0 To HW5DataSet.Tables(tableName).Rows.Count - 1
            Dim myPeriod As New Period
            myPeriod.ID = HW5DataSet.Tables(tableName).Rows(rowIndex)("ID")
            myPeriod.Length = HW5DataSet.Tables(tableName).Rows(rowIndex)("Length")
            myPeriod.PeriodName = HW5DataSet.Tables(tableName).Rows(rowIndex)("Periodname")
            Period.PeriodList.Add(myPeriod)
        Next
        '
        'RDB: We set properties of the Shift objects from data in the ShiftData table
        tableName = "ShiftData"
        HW5SQL = "SELECT * FROM " & tableName                               'RDB: This SQL selects all columns from the Meal table in MS Access
        HW5Data.RunSQL(tableName, HW5SQL, HW5DataSet, HW5ConnectionString)  'RDB: We execute the query
        '
        'RDB: We create the objects of the Employee Class from the data in the EmployeeData table
        For rowIndex As Integer = 0 To HW5DataSet.Tables(tableName).Rows.Count - 1
            Dim myOverlap As New Shift
            myOverlap.ID = HW5DataSet.Tables(tableName).Rows(rowIndex)("ID")
            myOverlap.ShiftName = HW5DataSet.Tables(tableName).Rows(rowIndex)("ShiftName")
            myOverlap.ShiftLength = HW5DataSet.Tables(tableName).Rows(rowIndex)("ShiftLength")
            For periodIndex = 0 To Period.PeriodList.Count - 1
                myOverlap.PeriodOverlap.Add(HW5DataSet.Tables(tableName).Rows(rowIndex)("Period Overlap " & CStr(periodIndex + 1)))
            Next
            Shift.ShiftList.Add(myOverlap)
            Next
            '
            'RDB: We set properties of the Demand objects from data in the Demand table
            tableName = "DemandData"
        HW5SQL = "SELECT * FROM " & tableName                               'RDB: This SQL selects all columns from the Meal table in MS Access
        HW5Data.RunSQL(tableName, HW5SQL, HW5DataSet, HW5ConnectionString)  'RDB: We execute the query
        '
        'RDB: We create the objects of the Demand Class from the data in the Meal table
        For rowIndex As Integer = 0 To HW5DataSet.Tables(tableName).Rows.Count - 1
            Dim myDemand As New Demand
            myDemand.ID = HW5DataSet.Tables(tableName).Rows(rowIndex)("ID")
            myDemand.DemandScenario = HW5DataSet.Tables(tableName).Rows(rowIndex)("DemandScenario")
            For periodIndex = 0 To Period.PeriodList.Count - 1
                myDemand.Period.Add(HW5DataSet.Tables(tableName).Rows(rowIndex)("Period " & CStr(periodIndex + 1)))
            Next
            Demand.DemandList.Add(myDemand)
            Next
            '
            'RDB: We set properties of the Employee objects from data in the EmployeeData table
            tableName = "EmployeeData"
        HW5SQL = "SELECT * FROM " & tableName                               'RDB: This SQL selects all columns from the Meal table in MS Access
        HW5Data.RunSQL(tableName, HW5SQL, HW5DataSet, HW5ConnectionString)  'RDB: We execute the query
        '
        'RDB: We create the objects of the Employee Class from the data in the Meal table
        For rowIndex As Integer = 0 To HW5DataSet.Tables(tableName).Rows.Count - 1
            Dim myEmployee As New Employee
            myEmployee.ID = HW5DataSet.Tables(tableName).Rows(rowIndex)("ID")
            myEmployee.EmployeeName = HW5DataSet.Tables(tableName).Rows(rowIndex)("Personnel")
            myEmployee.WageRate = HW5DataSet.Tables(tableName).Rows(rowIndex)("Wage Rate")
            For shiftIndex = 0 To Shift.ShiftList.Count - 1
                myEmployee.ShiftAvailability.Add(HW5DataSet.Tables(tableName).Rows(rowIndex)("Shift " & CStr(shiftIndex + 1)))
            Next
            Employee.EmployeeList.Add(myEmployee)
            Next

            '
            'RDB: We set properties of the ProcessData objects from data in the ProcessData table
            tableName = "ProcessData"
        HW5SQL = "SELECT * FROM " & tableName                               'RDB: This SQL selects all columns from the Meal table in MS Access
        HW5Data.RunSQL(tableName, HW5SQL, HW5DataSet, HW5ConnectionString)  'RDB: We execute the query
        '
        'RDB: We create the objects of the Process Class from the data in the ProcessData table
        For rowIndex As Integer = 0 To HW5DataSet.Tables(tableName).Rows.Count - 1
            Dim myProcessData As New Process
            myProcessData.ID = HW5DataSet.Tables(tableName).Rows(rowIndex)("ID")
            myProcessData.UM = HW5DataSet.Tables(tableName).Rows(rowIndex)("U/M")
            myProcessData.ProcessTime = HW5DataSet.Tables(tableName).Rows(rowIndex)("CycleTime")
            Process.ProcessList.Add(myProcessData)
        Next

    End Sub
End Class
