<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRDBMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.btnRDBOptimize = New System.Windows.Forms.Button()
        Me.HW5_SP19DataSet = New HW5_SP19.HW5_SP19DataSet()
        Me.HW5SP19DataSetBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        CType(Me.HW5_SP19DataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.HW5SP19DataSetBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnRDBOptimize
        '
        Me.btnRDBOptimize.BackColor = System.Drawing.Color.Snow
        Me.btnRDBOptimize.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRDBOptimize.Location = New System.Drawing.Point(96, 88)
        Me.btnRDBOptimize.Name = "btnRDBOptimize"
        Me.btnRDBOptimize.Size = New System.Drawing.Size(268, 54)
        Me.btnRDBOptimize.TabIndex = 0
        Me.btnRDBOptimize.Text = "Optimize My Schedule"
        Me.btnRDBOptimize.UseVisualStyleBackColor = False
        '
        'HW5_SP19DataSet
        '
        Me.HW5_SP19DataSet.DataSetName = "HW5_SP19DataSet"
        Me.HW5_SP19DataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'HW5SP19DataSetBindingSource
        '
        Me.HW5SP19DataSetBindingSource.DataSource = Me.HW5_SP19DataSet
        Me.HW5SP19DataSetBindingSource.Position = 0
        '
        'frmRDBMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.IndianRed
        Me.ClientSize = New System.Drawing.Size(465, 232)
        Me.Controls.Add(Me.btnRDBOptimize)
        Me.Name = "frmRDBMain"
        Me.Text = "Scheduler"
        CType(Me.HW5_SP19DataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.HW5SP19DataSetBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents btnRDBOptimize As Button
    Friend WithEvents HW5SP19DataSetBindingSource As BindingSource
    Friend WithEvents HW5_SP19DataSet As HW5_SP19DataSet
End Class
