<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Configurazioni
    Inherits System.Windows.Forms.Form

    'Form esegue l'override del metodo Dispose per pulire l'elenco dei componenti.
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

    'Richiesto da Progettazione Windows Form
    Private components As System.ComponentModel.IContainer

    'NOTA: la procedura che segue è richiesta da Progettazione Windows Form
    'Può essere modificata in Progettazione Windows Form.  
    'Non modificarla mediante l'editor del codice.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Configurazioni))
        Me.DataGridConfigurazioni = New System.Windows.Forms.DataGridView()
        Me.BTNSaveConfig = New System.Windows.Forms.Button()
        Me.BTNClose = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        CType(Me.DataGridConfigurazioni, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'DataGridConfigurazioni
        '
        Me.DataGridConfigurazioni.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridConfigurazioni.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridConfigurazioni.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridConfigurazioni.Location = New System.Drawing.Point(0, 0)
        Me.DataGridConfigurazioni.Name = "DataGridConfigurazioni"
        Me.DataGridConfigurazioni.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.DataGridConfigurazioni.Size = New System.Drawing.Size(843, 404)
        Me.DataGridConfigurazioni.TabIndex = 0
        '
        'BTNSaveConfig
        '
        Me.BTNSaveConfig.Location = New System.Drawing.Point(12, 410)
        Me.BTNSaveConfig.Name = "BTNSaveConfig"
        Me.BTNSaveConfig.Size = New System.Drawing.Size(75, 23)
        Me.BTNSaveConfig.TabIndex = 1
        Me.BTNSaveConfig.Text = "Salva"
        Me.BTNSaveConfig.UseVisualStyleBackColor = True
        '
        'BTNClose
        '
        Me.BTNClose.Location = New System.Drawing.Point(100, 410)
        Me.BTNClose.Name = "BTNClose"
        Me.BTNClose.Size = New System.Drawing.Size(75, 23)
        Me.BTNClose.TabIndex = 2
        Me.BTNClose.Text = "Chiudi"
        Me.BTNClose.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.DataGridConfigurazioni)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(843, 404)
        Me.Panel1.TabIndex = 3
        '
        'Configurazioni
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(843, 446)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.BTNClose)
        Me.Controls.Add(Me.BTNSaveConfig)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Configurazioni"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Configurazioni"
        CType(Me.DataGridConfigurazioni, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents DataGridConfigurazioni As DataGridView
    Friend WithEvents BTNSaveConfig As Button
    Friend WithEvents BTNClose As Button
    Friend WithEvents Panel1 As Panel
End Class
