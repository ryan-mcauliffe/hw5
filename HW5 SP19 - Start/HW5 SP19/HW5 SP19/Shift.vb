Public Class Shift
    Public Shared Property ShiftList As New List(Of Shift)

    Public Property ID As Long
    Public Property ShiftName As String
    Public Property ShiftLength As Single
    Public Property PeriodOverlap As New List(Of Single)
    '******************************************************************************
    Public Sub New()

    End Sub

End Class
