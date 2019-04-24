Public Class Employee
    Public Shared Property EmployeeList As New List(Of Employee)

    Public Property ID As Long
    Public Property EmployeeName As String
    Public Property ShiftAvailability As New List(Of Integer)
    Public Property WageRate As Single
    '******************************************************************************
    '******************************************************************************
    Public Sub New()

    End Sub

End Class
