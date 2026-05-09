<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class InterfaceWindow
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.OutputBox = New System.Windows.Forms.TextBox()
        Me.CommandBox = New System.Windows.Forms.TextBox()
        Me.EnterButton = New System.Windows.Forms.Button()
        Me.ToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.StatusBar = New System.Windows.Forms.StatusStrip()
        Me.CPUActiveLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusBar.SuspendLayout()
        Me.SuspendLayout()
        '
        'OutputBox
        '
        Me.OutputBox.AllowDrop = True
        Me.OutputBox.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.OutputBox.BackColor = System.Drawing.SystemColors.Window
        Me.OutputBox.Font = New System.Drawing.Font("Consolas", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OutputBox.Location = New System.Drawing.Point(-1, 2)
        Me.OutputBox.Multiline = True
        Me.OutputBox.Name = "OutputBox"
        Me.OutputBox.ReadOnly = True
        Me.OutputBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.OutputBox.Size = New System.Drawing.Size(800, 392)
        Me.OutputBox.TabIndex = 2
        Me.OutputBox.TabStop = False
        '
        'CommandBox
        '
        Me.CommandBox.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CommandBox.Font = New System.Drawing.Font("Consolas", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CommandBox.Location = New System.Drawing.Point(12, 400)
        Me.CommandBox.Name = "CommandBox"
        Me.CommandBox.Size = New System.Drawing.Size(688, 26)
        Me.CommandBox.TabIndex = 0
        '
        'EnterButton
        '
        Me.EnterButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.EnterButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.EnterButton.Location = New System.Drawing.Point(713, 402)
        Me.EnterButton.Name = "EnterButton"
        Me.EnterButton.Size = New System.Drawing.Size(75, 23)
        Me.EnterButton.TabIndex = 1
        Me.EnterButton.Text = "&Enter"
        Me.EnterButton.UseVisualStyleBackColor = True
        '
        'StatusBar
        '
        Me.StatusBar.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CPUActiveLabel})
        Me.StatusBar.Location = New System.Drawing.Point(0, 435)
        Me.StatusBar.Name = "StatusBar"
        Me.StatusBar.Size = New System.Drawing.Size(800, 22)
        Me.StatusBar.TabIndex = 3
        '
        'CPUActiveLabel
        '
        Me.CPUActiveLabel.Name = "CPUActiveLabel"
        Me.CPUActiveLabel.Size = New System.Drawing.Size(0, 17)
        '
        'InterfaceWindow
        '
        Me.AcceptButton = Me.EnterButton
        Me.AllowDrop = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 457)
        Me.Controls.Add(Me.StatusBar)
        Me.Controls.Add(Me.EnterButton)
        Me.Controls.Add(Me.CommandBox)
        Me.Controls.Add(Me.OutputBox)
        Me.KeyPreview = True
        Me.Name = "InterfaceWindow"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.StatusBar.ResumeLayout(False)
        Me.StatusBar.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents OutputBox As System.Windows.Forms.TextBox
    Friend WithEvents CommandBox As System.Windows.Forms.TextBox
    Friend WithEvents EnterButton As System.Windows.Forms.Button
    Friend WithEvents ToolTip As System.Windows.Forms.ToolTip
    Friend WithEvents StatusBar As System.Windows.Forms.StatusStrip
    Friend WithEvents CPUActiveLabel As System.Windows.Forms.ToolStripStatusLabel
End Class
