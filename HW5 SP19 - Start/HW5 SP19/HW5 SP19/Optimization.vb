Imports Microsoft.SolverFoundation.Common
Imports Microsoft.SolverFoundation.Services
Imports Microsoft.SolverFoundation.Solvers
'****************************************************************************************************************
Public Class Optimization
    Dim HW5RPMModel As New SimplexSolver
    Dim dvKey As String
    Dim dvindex As Integer
    Dim coefficient As Single
    Dim constraintKey As String
    Dim ConstraintIndex As Integer
    Dim objKey As String = "objective Function"

    Dim objIndex As Integer
    Public optimalObj As Single
    Public dvValue(Period.PeriodList.Count - 1, Process.ProcessList.Count - 1) As Single

    Public Shared ReadOnly NegativeInfinity As Rational
    Public Shared ReadOnly PositiveInfinity As Rational
    '****************************************************************************************************************
    Public Sub RPMBuildModel()

        'Decision Variables
        For Each myShift As Shift In Shift.ShiftList
            For Each myEmployee As Employee In Employee.EmployeeList
                dvKey = myShift.ShiftName & "_" & myEmployee.EmployeeName
                HW5RPMModel.AddVariable(dvKey, dvindex)
                HW5RPMModel.SetBounds(dvindex, 0, Rational.PositiveInfinity)
            Next
        Next

        '************************************************************************************************************
        'Constraints
        '
        'Capacity Constraints
        For Each myEmployee As Employee In Employee.EmployeeList
            constraintKey = "Capacity Constraint" & "_" & myEmployee.EmployeeName
            HW5RPMModel.AddRow(constraintKey, ConstraintIndex)

            For Each shifty As Shift In Shift.ShiftList
                dvindex = HW5RPMModel.GetIndexFromKey(dvKey)
                coefficient = 1
                HW5RPMModel.SetCoefficient(ConstraintIndex, dvindex, coefficient)
            Next
            HW5RPMModel.SetBounds(ConstraintIndex, 0, )
        Next

        'Requirement Constraints
        For Each shifty As Shift In Shift.ShiftList
            constraintKey = "Capacity Constraint" & "_" & shifty.ShiftName
            HW5RPMModel.AddRow(constraintKey, ConstraintIndex)

            For Each employ As Employee In Employee.EmployeeList
                dvindex = HW5RPMModel.GetIndexFromKey(dvKey)
                coefficient = 1
                HW5RPMModel.SetCoefficient(ConstraintIndex, dvindex, coefficient)
            Next

            HW5RPMModel.SetBounds(ConstraintIndex, 0, shifty.ShiftLength)
        Next
        '************************************************************************************************************
        '*************************************************************************************************************
        'ObjectiveFunction
        HW5RPMModel.AddRow(objKey, objIndex)
        For Each myShift As Shift In Shift.ShiftList
            For Each emp As Employee In Employee.EmployeeList
                dvindex = HW5RPMModel.GetIndexFromKey(myShift.ShiftName & "_" & emp.EmployeeName)
                'An alternate way to get the coefficents
                'coefficient = effect.Effect(emp.ActivityList.IndexOf(activity))

                If emp.EmployeeName = "Employee 1" Then coefficient = 1
                If emp.EmployeeName = "Employee 2" Then coefficient = 2
                If emp.EmployeeName = "Employee 3" Then coefficient = 3
                If emp.EmployeeName = "Employee 4" Then coefficient = 4
                HW5RPMModel.SetCoefficient(objIndex, dvindex, coefficient)
            Next
        Next
        HW5RPMModel.AddGoal(objIndex, 0, False)

    End Sub
    '****************************************************************************************************************
    Public Sub RPMRunModel()
        '----------------------------------------------------------------------------------------------------------
        'RDB:  Solve the optimization

        Dim mySolverParms As New SimplexSolverParams
        mySolverParms.MixedIntegerGapTolerance = 1              'RDB: For IP only - 1 percent gap tolerance between upper and lower bounds of objective function
        mySolverParms.VariableFeasibilityTolerance = 0.00001    'RDB: For IP only - required closeness to a whole number of each variable
        mySolverParms.MaxPivotCount = 100000                     'RDB: Number of iterations.  Increase as necessary
        HW5RPMModel.Solve(mySolverParms)

        'RDB: We check to see if we got an answer
        If HW5RPMModel.Result = LinearResult.UnboundedPrimal Then
            MessageBox.Show("Solution is unbounded")
            Exit Sub
        ElseIf _
        HW5RPMModel.Result = LinearResult.InfeasiblePrimal Then
            MessageBox.Show("Decision model is infeasible")
            Exit Sub
        Else
            ShowAnswer()
        End If
    End Sub
    '***************************************************************************************************************
    Public Sub ShowAnswer()
        '----------------------------------------------------------------------------------------------------------
        'RDB: Now we display the optimal values of the variables and objective function
        optimalObj = CSng(HW5RPMModel.GetValue(objIndex).ToDouble)

        'RDB: We transfer the values of the decision variables to an array 
        Dim rowIndex As Integer = 0
        Dim columnIndex As Integer = 0

        '
        For Each proc As Process In Process.ProcessList
            rowIndex = Process.ProcessList.IndexOf(proc)
            For Each shifty As Shift In Shift.ShiftList
                columnIndex = Shift.ShiftList.IndexOf(shifty)
                dvKey = proc.ProcessTime & "_" & shifty.ShiftLength
                dvindex = HW5RPMModel.GetIndexFromKey(dvKey)
                dvValue(rowIndex, columnIndex) = CSng(HW5RPMModel.GetValue(dvindex).ToDouble)
            Next
        Next
        '************************************************************************************
        Solution.RPMTable.CellBorderStyle = TableLayoutPanelCellBorderStyle.Single
        '
        'RDB: We enter the column headings into the table
        For column As Integer = 1 To Solution.RPMTable.ColumnCount - 1
            Dim myLabel As New Label
            myLabel.Text = "Activity " & CStr(column)
            Solution.RPMTable.Controls.Add(myLabel)
            myLabel.Visible = True
            myLabel.TextAlign = ContentAlignment.MiddleCenter
            Solution.RPMTable.SetRow(myLabel, 0)
            Solution.RPMTable.SetColumn(myLabel, column)
            myLabel.Anchor = AnchorStyles.Bottom
            myLabel.Anchor = AnchorStyles.Top
            myLabel.Anchor = AnchorStyles.Left
            myLabel.Anchor = AnchorStyles.Right

        Next
        '
        'RDB: We enter the row headings into the table
        rowIndex = 0
        For Each proc As Process In Process.ProcessList
            Dim myLabel As New Label
            myLabel.Text = proc.ProcessTime
            myLabel.Visible = True
            myLabel.TextAlign = ContentAlignment.MiddleCenter
            Solution.RPMTable.SetRow(myLabel, rowIndex + 1)
            Solution.RPMTable.SetColumn(myLabel, 0)
            Solution.RPMTable.Dock = DockStyle.Fill
            Solution.RPMTable.Controls.Add(myLabel)
            myLabel.Anchor = AnchorStyles.Bottom
            myLabel.Anchor = AnchorStyles.Top
            myLabel.Anchor = AnchorStyles.Left
            myLabel.Anchor = AnchorStyles.Right
            rowIndex += 1
        Next

        For row As Integer = 1 To Solution.RPMTable.RowCount - 1
            For column As Integer = 1 To Solution.RPMTable.ColumnCount - 1
                Dim myLabel As New Label
                myLabel.Text = CStr(dvValue(row - 1, column - 1))
                myLabel.Visible = True
                myLabel.TextAlign = ContentAlignment.MiddleCenter
                Solution.RPMTable.SetRow(myLabel, row)
                Solution.RPMTable.SetColumn(myLabel, column)
                Solution.RPMTable.Dock = DockStyle.Fill
                Solution.RPMTable.Controls.Add(myLabel)
                myLabel.Anchor = AnchorStyles.Bottom
                myLabel.Anchor = AnchorStyles.Top
                myLabel.Anchor = AnchorStyles.Left
                myLabel.Anchor = AnchorStyles.Right
            Next
        Next
        Solution.Show()

    End Sub
End Class

