<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_Zoom
    Inherits System.Windows.Forms.Form

    'Form esegue l'override del metodo Dispose per pulire l'elenco dei componenti.
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

    'Richiesto da Progettazione Windows Form
    Private components As System.ComponentModel.IContainer

    'NOTA: la procedura che segue è richiesta da Progettazione Windows Form
    'Può essere modificata in Progettazione Windows Form.  
    'Non modificarla mediante l'editor del codice.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Picture_Zoom = New System.Windows.Forms.PictureBox()
        CType(Me.Picture_Zoom, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Picture_Zoom
        '
        Me.Picture_Zoom.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Picture_Zoom.Location = New System.Drawing.Point(0, 0)
        Me.Picture_Zoom.Name = "Picture_Zoom"
        Me.Picture_Zoom.Size = New System.Drawing.Size(800, 450)
        Me.Picture_Zoom.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.Picture_Zoom.TabIndex = 1
        Me.Picture_Zoom.TabStop = False
        '
        'Form_Zoom
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.Picture_Zoom)
        Me.Name = "Form_Zoom"
        Me.Text = "Form_Zoom"
        CType(Me.Picture_Zoom, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Picture_Zoom As PictureBox
End Class
