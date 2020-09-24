<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmImportMain
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
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

    'Wird vom Windows Form-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Windows Form-Designer erforderlich.
    'Das Bearbeiten ist mit dem Windows Form-Designer möglich.  
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.cmbBuha = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.butKreditoren = New System.Windows.Forms.Button()
        Me.butDebitoren = New System.Windows.Forms.Button()
        Me.dgvDebitoren = New System.Windows.Forms.DataGridView()
        Me.butImport = New System.Windows.Forms.Button()
        Me.dgvDebitorenSub = New System.Windows.Forms.DataGridView()
        Me.dgvInfo = New System.Windows.Forms.DataGridView()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgvDebitoren, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvDebitorenSub, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvInfo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmbBuha
        '
        Me.cmbBuha.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbBuha.FormattingEnabled = True
        Me.cmbBuha.Location = New System.Drawing.Point(66, 24)
        Me.cmbBuha.Name = "cmbBuha"
        Me.cmbBuha.Size = New System.Drawing.Size(270, 28)
        Me.cmbBuha.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 27)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(47, 20)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Buha"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.butKreditoren)
        Me.GroupBox1.Controls.Add(Me.butDebitoren)
        Me.GroupBox1.Location = New System.Drawing.Point(66, 66)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(281, 68)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Modus"
        '
        'butKreditoren
        '
        Me.butKreditoren.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.butKreditoren.Location = New System.Drawing.Point(139, 19)
        Me.butKreditoren.Name = "butKreditoren"
        Me.butKreditoren.Size = New System.Drawing.Size(131, 41)
        Me.butKreditoren.TabIndex = 1
        Me.butKreditoren.Text = "&Kredtiroren"
        Me.butKreditoren.UseVisualStyleBackColor = True
        '
        'butDebitoren
        '
        Me.butDebitoren.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.butDebitoren.Location = New System.Drawing.Point(16, 19)
        Me.butDebitoren.Name = "butDebitoren"
        Me.butDebitoren.Size = New System.Drawing.Size(117, 41)
        Me.butDebitoren.TabIndex = 0
        Me.butDebitoren.Text = "&Debitoren"
        Me.butDebitoren.UseVisualStyleBackColor = True
        '
        'dgvDebitoren
        '
        Me.dgvDebitoren.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvDebitoren.Location = New System.Drawing.Point(12, 158)
        Me.dgvDebitoren.Name = "dgvDebitoren"
        Me.dgvDebitoren.Size = New System.Drawing.Size(1613, 467)
        Me.dgvDebitoren.TabIndex = 3
        '
        'butImport
        '
        Me.butImport.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.butImport.Location = New System.Drawing.Point(1493, 84)
        Me.butImport.Name = "butImport"
        Me.butImport.Size = New System.Drawing.Size(132, 41)
        Me.butImport.TabIndex = 4
        Me.butImport.Text = "&Import"
        Me.butImport.UseVisualStyleBackColor = True
        '
        'dgvDebitorenSub
        '
        Me.dgvDebitorenSub.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvDebitorenSub.Location = New System.Drawing.Point(353, 19)
        Me.dgvDebitorenSub.Name = "dgvDebitorenSub"
        Me.dgvDebitorenSub.Size = New System.Drawing.Size(774, 133)
        Me.dgvDebitorenSub.TabIndex = 5
        '
        'dgvInfo
        '
        Me.dgvInfo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvInfo.DefaultCellStyle = DataGridViewCellStyle1
        Me.dgvInfo.Location = New System.Drawing.Point(1133, 19)
        Me.dgvInfo.Name = "dgvInfo"
        Me.dgvInfo.Size = New System.Drawing.Size(354, 133)
        Me.dgvInfo.TabIndex = 7
        '
        'frmImportMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1637, 637)
        Me.Controls.Add(Me.dgvInfo)
        Me.Controls.Add(Me.dgvDebitorenSub)
        Me.Controls.Add(Me.butImport)
        Me.Controls.Add(Me.dgvDebitoren)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmbBuha)
        Me.Name = "frmImportMain"
        Me.Text = "Sage200 - Importer"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.dgvDebitoren, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvDebitorenSub, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvInfo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents cmbBuha As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents butKreditoren As Button
    Friend WithEvents butDebitoren As Button
    Friend WithEvents dgvDebitoren As DataGridView
    Friend WithEvents butImport As Button
    Friend WithEvents dgvDebitorenSub As DataGridView
    Friend WithEvents dgvInfo As DataGridView
End Class
