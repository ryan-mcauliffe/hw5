Public Class frmRDBMain
    Private Sub btnRDBOptimize_Click(sender As Object, e As EventArgs) Handles btnRDBOptimize.Click
        'RPM: We create an instance of the CreateObjects class and run the method to import data and create objects and lists of objects
        Dim myObjectCreator As New ObjectCreator
        myObjectCreator.CreateObjectsAndLists()
        '
        'RPM: Now we create an Optimization object to build an run the linear program
        Dim myOptimization As New Optimization
        myOptimization.RPMBuildModel()
        myOptimization.RPMRunModel()


    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class