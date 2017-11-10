<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class TimboxForm
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
        Me.Timbrar = New System.Windows.Forms.Button()
        Me.Cancelar = New System.Windows.Forms.Button()
        Me.uuid_label = New System.Windows.Forms.Label()
        Me.txtUUID = New System.Windows.Forms.TextBox()
        Me.responseBox = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'Timbrar
        '
        Me.Timbrar.Location = New System.Drawing.Point(67, 41)
        Me.Timbrar.Name = "Timbrar"
        Me.Timbrar.Size = New System.Drawing.Size(100, 33)
        Me.Timbrar.TabIndex = 0
        Me.Timbrar.Text = "Timbrar"
        Me.Timbrar.UseVisualStyleBackColor = True
        '
        'Cancelar
        '
        Me.Cancelar.Location = New System.Drawing.Point(67, 100)
        Me.Cancelar.Name = "Cancelar"
        Me.Cancelar.Size = New System.Drawing.Size(100, 33)
        Me.Cancelar.TabIndex = 1
        Me.Cancelar.Text = "Cancelar"
        Me.Cancelar.UseVisualStyleBackColor = True
        '
        'uuid_label
        '
        Me.uuid_label.AutoSize = True
        Me.uuid_label.Location = New System.Drawing.Point(233, 61)
        Me.uuid_label.Name = "uuid_label"
        Me.uuid_label.Size = New System.Drawing.Size(87, 13)
        Me.uuid_label.TabIndex = 2
        Me.uuid_label.Text = "UUID a cancelar"
        '
        'txtUUID
        '
        Me.txtUUID.Location = New System.Drawing.Point(236, 77)
        Me.txtUUID.Name = "txtUUID"
        Me.txtUUID.Size = New System.Drawing.Size(221, 20)
        Me.txtUUID.TabIndex = 3
        '
        'responseBox
        '
        Me.responseBox.Location = New System.Drawing.Point(67, 153)
        Me.responseBox.Multiline = True
        Me.responseBox.Name = "responseBox"
        Me.responseBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.responseBox.Size = New System.Drawing.Size(552, 206)
        Me.responseBox.TabIndex = 4
        '
        'TimboxForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(683, 394)
        Me.Controls.Add(Me.responseBox)
        Me.Controls.Add(Me.txtUUID)
        Me.Controls.Add(Me.uuid_label)
        Me.Controls.Add(Me.Cancelar)
        Me.Controls.Add(Me.Timbrar)
        Me.Name = "TimboxForm"
        Me.Text = "Timbox Integracion-VB"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Timbrar As Button
    Friend WithEvents Cancelar As Button
    Friend WithEvents uuid_label As Label
    Friend WithEvents txtUUID As TextBox
    Friend WithEvents responseBox As TextBox
End Class
