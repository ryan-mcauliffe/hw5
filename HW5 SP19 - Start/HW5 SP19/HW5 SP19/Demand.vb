Public Class Demand
    Public Shared Property DemandList As New List(Of Demand)

    Public Property ID As Long
    Public Property DemandScenario As String
    Public Property Period As New List(Of Integer)
    '******************************************************************************
    Public Sub New()

    End Sub

End Class
